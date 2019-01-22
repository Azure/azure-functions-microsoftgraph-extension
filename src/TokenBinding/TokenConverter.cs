// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;

    public class TokenConverter : IAsyncConverter<TokenAttribute, string>, IAsyncConverter<TokenBaseAttribute, string>
    {
        private IOptions<TokenOptions> _options;
        private IAadClient _aadManager;
        private IEasyAuthClient _easyAuthClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="Converters"/> class.
        /// </summary>
        /// <param name="parent">TokenExtensionConfig containing necessary context & methods</param>
        public TokenConverter(IOptions<TokenOptions> options, IEasyAuthClient easyAuthClient, IAadClient aadClient)
        {
            _options = options;
            _easyAuthClient = easyAuthClient;
            _aadManager = aadClient;
        }

        public async Task<string> ConvertAsync(TokenBaseAttribute attribute, CancellationToken cancellationToken)
        {
            attribute.CheckValidity();
            switch (attribute.Identity)
            {
                case TokenIdentityMode.UserFromId:
                    // If the attribute has no identity provider, assume AAD
                    attribute.IdentityProvider = attribute.IdentityProvider ?? "AAD";
                    var easyAuthTokenManager = new EasyAuthTokenManager(_easyAuthClient, _options);
                    return await easyAuthTokenManager.GetEasyAuthAccessTokenAsync(attribute);
                case TokenIdentityMode.UserFromToken:
                    return await GetAuthTokenFromUserToken(attribute.UserToken, attribute.Resource);
                case TokenIdentityMode.ClientCredentials:
                    return await _aadManager.GetTokenFromClientCredentials(attribute.Resource);
                case TokenIdentityMode.AppIdentity:
                    return await _aadManager.GetTokenFromAppIdentity(attribute.Resource, attribute.ConnectionString);
            }

            throw new InvalidOperationException("Unable to authorize without Principal ID or ID Token.");
        }

        public async Task<string> ConvertAsync(TokenAttribute attribute, CancellationToken cancellationToken)
        {
            return await ConvertAsync(attribute as TokenBaseAttribute, cancellationToken);
        }

        private async Task<string> GetAuthTokenFromUserToken(string userToken, string resource)
        {
            if (string.IsNullOrWhiteSpace(resource))
            {
                throw new ArgumentException("A resource is required to get an auth token on behalf of a user.");
            }

            // If the incoming token already has the correct audience (resource), then skip the exchange (it will fail with AADSTS50013!)
            var currentAudience = GetAudience(userToken);
            if (currentAudience != resource)
            {
                string token = await _aadManager.GetTokenOnBehalfOfUserAsync(
                    userToken,
                    resource);
                return token;
            }

            // No exchange requested, return token directly.
            return userToken;
        }

        private string GetAudience(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var audience = jwt.Audiences.FirstOrDefault();
            return audience;
        }
    }
}
