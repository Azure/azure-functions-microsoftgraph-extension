using System;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    internal class TokenConverter : IAsyncConverter<TokenAttribute, string>
    {
        private IOptions<TokenOptions> _options;
        private IAadClient _aadManager;
        private IEasyAuthClient _easyAuthClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="Converters"/> class.
        /// </summary>
        /// <param name="parent">TokenExtensionConfig containing necessary context & methods</param>
        internal TokenConverter(IOptions<TokenOptions> options, IEasyAuthClient easyAuthClient, IAadClient aadClient)
        {
            _options = options;
            _easyAuthClient = easyAuthClient;
            _aadManager = aadClient;
        }

        /// <summary>
        /// Convert from a TokenAttribute to a string (to use as input to a fx)
        /// </summary>
        /// <param name="input">TokenAttribute with necessary user info & desired resource</param>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>JWT</returns>
        public async Task<string> ConvertAsync(TokenAttribute attribute, CancellationToken cancellationToken)
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
            }

            throw new InvalidOperationException("Unable to authorize without Principal ID or ID Token.");
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
