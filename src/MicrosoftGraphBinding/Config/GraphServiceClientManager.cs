// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System;
    using System.Collections.Concurrent;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Graph;

    internal class GraphServiceClientManager
    {
        private readonly IAsyncConverter<TokenBaseAttribute, string> _tokenProvider;
        private readonly IGraphServiceClientProvider _clientProvider;
        private readonly GraphOptions _options;

        /// <summary>
        /// Map principal Id + scopes -> GraphServiceClient + token expiration date
        /// </summary>
        private ConcurrentDictionary<string, CachedClient> _clients = new ConcurrentDictionary<string, CachedClient>();

        public GraphServiceClientManager(GraphOptions options, IAsyncConverter<TokenBaseAttribute, string> tokenProvider, IGraphServiceClientProvider clientProvider)
        {
            _tokenProvider = tokenProvider;
            _clientProvider = clientProvider;
            _options = options;
        }

        /// <summary>
        /// Retrieve audience from raw JWT
        /// </summary>
        /// <param name="rawToken">JWT</param>
        /// <returns>Token audience</returns>
        private static string GetTokenOID(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var oidClaim = jwt.Claims.FirstOrDefault(claim => claim.Type == "oid");
            if (oidClaim == null)
            {
                throw new InvalidOperationException("The graph token is missing an oid. Check your Microsoft Graph binding configuration.");
            }
            return oidClaim.Value;
        }

        /// <summary>
        /// Given a JWT, return the list of scopes in alphabetical order
        /// </summary>
        /// <param name="rawToken">raw JWT</param>
        /// <returns>string of scopes in alphabetical order, separated by a space</returns>
        private static string GetTokenOrderedScopes(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var stringScopes = jwt.Claims.FirstOrDefault(claim => claim.Type == "scp")?.Value;
            if (stringScopes == null)
            {
                throw new InvalidOperationException("The graph token has no scopes. Ensure your application is properly configured to access the Microsoft Graph.");
            }
            var scopes = stringScopes.Split(' ');
            Array.Sort(scopes);
            return string.Join(" ", scopes);
        }

        /// <summary>
        /// Retrieve integer token expiration date
        /// </summary>
        /// <param name="rawToken">raw JWT</param>
        /// <returns>parsed expiration date</returns>
        private static int GetTokenExpirationDate(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var stringTime = jwt.Claims.FirstOrDefault(claim => claim.Type == "exp").Value;
            int result;
            if (int.TryParse(stringTime, out result))
            {
                return result;
            }
            else
            {
                return -1;
            }
        }

        /// <summary>
        /// Hydrate GraphServiceClient from a moniker (serialized TokenAttribute)
        /// </summary>
        /// <param name="moniker">string representing serialized TokenAttribute</param>
        /// <returns>Authenticated GraphServiceClient</returns>
        public async Task<IGraphServiceClient> GetMSGraphClientFromUserIdAsync(string userId, CancellationToken token)
        {
            var attr = new TokenAttribute
            {
                AadResource = _options.GraphBaseUrl,
                Identity = TokenIdentityMode.UserFromRequest,
            };

            return await this.GetMSGraphClientFromTokenAttributeAsync(attr, token);
        }

        /// <summary>
        /// Either retrieve existing GSC or create a new one
        /// GSCs are cached using a combination of the user's principal ID and the scopes of the token used to authenticate
        /// </summary>
        /// <param name="attribute">Token attribute with either principal ID or ID token</param>
        /// <returns>Authenticated GSC</returns>
        public virtual async Task<IGraphServiceClient> GetMSGraphClientFromTokenAttributeAsync(TokenBaseAttribute attribute, CancellationToken cancellationToken)
        {
            string token = await this._tokenProvider.ConvertAsync(attribute, cancellationToken);
            string principalId = GetTokenOID(token);

            var key = string.Concat(principalId, " ", GetTokenOrderedScopes(token));

            CachedClient cachedClient = null;

            // Check to see if there already exists a GSC associated with this principal ID and the token scopes.
            if (_clients.TryGetValue(key, out cachedClient))
            {
                // Check if token is expired
                if (cachedClient.expirationDate < DateTimeOffset.Now.ToUnixTimeSeconds())
                {
                    // Need to update the client's token & expiration date
                    // $$ todo -- just reset token instead of whole new authentication provider?
                    _clientProvider.UpdateGraphServiceClientAuthToken(cachedClient.client, token);
                    cachedClient.expirationDate = GetTokenExpirationDate(token);
                }

                return cachedClient.client;
            }
            else
            {
                cachedClient = new CachedClient
                {
                    client = _clientProvider.CreateNewGraphServiceClient(token),
                    expirationDate = GetTokenExpirationDate(token),
                };
                _clients.TryAdd(key, cachedClient);
                return cachedClient.client;
            }
        }
    }
}
