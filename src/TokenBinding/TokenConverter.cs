// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;

    public class TokenConverter : IAsyncConverter<TokenAttribute, string>, IAsyncConverter<TokenBaseAttribute, string>
    {
        private TokenOptions _options;
        private IEasyAuthClient _easyAuthClient;

        // Lazily initialize as we don't want to force validation of parameters unless we are actually using them.
        private Lazy<IAadService> _aadService;

        /// <summary>
        /// Initializes a new instance of the <see cref="Converters"/> class.
        /// </summary>
        /// <param name="parent">TokenExtensionConfig containing necessary context & methods</param>
        public TokenConverter(TokenOptions options, IEasyAuthClient easyAuthClient, IAadServiceFactory aadServiceFactory)
        {
            _options = options;
            _easyAuthClient = easyAuthClient;
            _aadService = new Lazy<IAadService>(() => aadServiceFactory.GetAadClient(_options.TenantUrl, _options.ClientId, _options.ClientSecret));
        }

        public async Task<string> ConvertAsync(TokenBaseAttribute attribute, CancellationToken cancellationToken)
        {
            switch (attribute.Identity)
            {
                case TokenIdentityMode.UserFromRequest:
                    return await GetResourceTokenFromAccessToken(attribute.EasyAuthAccessToken, attribute.AadResource);
                case TokenIdentityMode.ClientCredentials:
                    return await _aadService.Value.GetTokenFromClientCredentials(attribute.AadResource);
            }

            throw new InvalidOperationException("Unable to authorize without Principal ID or ID Token.");
        }

        public async Task<string> ConvertAsync(TokenAttribute attribute, CancellationToken cancellationToken)
        {
            return await ConvertAsync(attribute as TokenBaseAttribute, cancellationToken);
        }

        private async Task<string> GetResourceTokenFromAccessToken(string userToken, string resource)
        {
            if (string.IsNullOrWhiteSpace(resource))
            {
                throw new ArgumentException("A resource is required to get an auth token on behalf of a user.");
            }

            // If the incoming token already has the correct audience (resource), then skip the exchange (it will fail with AADSTS50013!)
            var currentAudience = GetAudience(userToken);
            if (currentAudience != resource)
            {
                string token = await _aadService.Value.GetTokenOnBehalfOfUserAsync(
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
