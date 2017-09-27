// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System;
    using System.Collections.Concurrent;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Graph;

    internal class ServiceManager
    {
        public IExcelClient ExcelClient { get; set; }
        public IOutlookClient OutlookClient { get; set; }
        public IOneDriveClient OneDriveClient { get; set; }

        internal AuthTokenExtensionConfig _tokenExtension;

        /// <summary>
        /// Map principal Id + scopes -> GraphServiceClient + token expiration date
        /// </summary>
        private ConcurrentDictionary<string, CachedClient> _clients = new ConcurrentDictionary<string, CachedClient>();

        public ServiceManager(AuthTokenExtensionConfig config)
        {
            _tokenExtension = config;
        }


        internal ExcelService GetExcelManager(TokenAttribute attribute)
        {
            return new ExcelService(GetExcelClient(attribute));
        }

        private IExcelClient GetExcelClient(TokenAttribute attribute)
        {
            return ExcelClient ?? new ExcelClient(GetMSGraphClientAsync(attribute));
        }

        internal OutlookService GetOutlookService(TokenAttribute attribute)
        {
            return new OutlookService(GetOutlookClient(attribute));
        }

        private IOutlookClient GetOutlookClient(TokenAttribute attribute)
        {
            return OutlookClient ?? new OutlookClient(GetMSGraphClientAsync(attribute));
        }

        internal OneDriveService GetOneDriveService(TokenAttribute attribute)
        {
            return new OneDriveService(GetOneDriveClient(attribute));
        }

        private IOneDriveClient GetOneDriveClient(TokenAttribute attribute)
        {
            return OneDriveClient ?? new OneDriveClient(GetMSGraphClientAsync(attribute));
        }

        /// <summary>
        /// Retrieve audience from raw JWT
        /// </summary>
        /// <param name="rawToken">JWT</param>
        /// <returns>Token audience</returns>
        public static string GetTokenOID(string rawToken)
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
        public static string GetTokenOrderedScopes(string rawToken)
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
        public static int GetTokenExpirationDate(string rawToken)
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
        internal async Task<IGraphServiceClient> GetMSGraphClientFromUserIdAsync(string userId)
        {
            var attr = new TokenAttribute
            {
                UserId = userId,
                Resource = O365Constants.GraphBaseUrl,
                Identity = TokenIdentityMode.UserFromId,
            };

            return await this.GetMSGraphClientAsync(attr);
        }

        /// <summary>
        /// Either retrieve existing GSC or create a new one
        /// GSCs are cached using a combination of the user's principal ID and the scopes of the token used to authenticate
        /// </summary>
        /// <param name="attribute">Token attribute with either principal ID or ID token</param>
        /// <returns>Authenticated GSC</returns>
        public async Task<IGraphServiceClient> GetMSGraphClientAsync(TokenAttribute attribute)
        {
            string token = await this._tokenExtension.GetAccessTokenAsync(attribute);
            string principalId = GetTokenOID(token);

            var key = string.Concat(principalId, " ", GetTokenOrderedScopes(token));

            CachedClient client = null;

            // Check to see if there already exists a GSC associated with this principal ID and the token scopes.
            if (this._clients.TryGetValue(key, out client))
            {
                // Check if token is expired
                if (client.expirationDate < DateTimeOffset.Now.ToUnixTimeSeconds())
                {
                    // Need to update the client's token & expiration date
                    // $$ todo -- just reset token instead of whole new authentication provider?
                    client.client.AuthenticationProvider = new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

                            return Task.CompletedTask;
                        });
                    client.expirationDate = GetTokenExpirationDate(token);
                }

                return client.client;
            }
            else
            {
                client = new CachedClient
                {
                    client = new GraphServiceClient(
                        new DelegateAuthenticationProvider(
                            (requestMessage) =>
                            {
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                                return Task.CompletedTask;
                            })),
                    expirationDate = GetTokenExpirationDate(token),
                };
                this._clients.TryAdd(key, client);
                return client.client;
            }
        }


    }
}
