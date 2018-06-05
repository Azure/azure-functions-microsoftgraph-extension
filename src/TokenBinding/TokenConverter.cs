// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.IdentityModel.Tokens;

    /// <summary>
    /// Class representing an application's [EasyAuth] Token Store
    /// see  https://cgillum.tech/2016/03/07/app-service-token-store/
    /// </summary>
    internal class TokenConverter : IAsyncConverter<TokenAttribute, string>
    {
        internal static readonly JwtSecurityTokenHandler JwtHandler = new JwtSecurityTokenHandler();

        private static readonly int GraphTokenBufferInMinutes = 5;

        private static readonly int _jwtTokenDurationInMinutes = 15;

        //Since the app setting storing the signing key is only used in some scenarios, initialize lazily
        private string SigningKey
        {
            get
            {
                if(_signingKey == null)
                {
                    _signingKey = _appSettings.Resolve(Constants.AppSettingWebsiteAuthSigningKey);
                }
                return _signingKey;
            }
        }

        private string _signingKey;

        private readonly IEasyAuthClient _easyAuthClient;

        private readonly IAadClient _aadClient;

        private readonly INameResolver _appSettings;

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenConverter"/> class.
        /// </summary>
        /// <param name="hostName">The hostname of the keystore. </param>
        /// <param name="signingKey">The website authorization signing key</param>
        public TokenConverter(IEasyAuthClient easyAuthClient, IAadClient aadClient, INameResolver appSettings)
        {
            _easyAuthClient = easyAuthClient;
            _aadClient = aadClient;
            _appSettings = appSettings;
        }

        public async Task<string> ConvertAsync(TokenAttribute input, CancellationToken cancellationToken)
        {
            input.CheckValidity();

            input.IdentityProvider = input.IdentityProvider ?? "AAD";
            switch (input.Identity)
            {
                case TokenIdentityMode.UserFromId:
                    return await GetEasyAuthAccessTokenAsync(input);
                case TokenIdentityMode.UserFromToken:
                    return await GetAuthTokenFromUserTokenAsync(input.UserToken, input.Resource);
                case TokenIdentityMode.ClientCredentials:
                    return await _aadClient.GetTokenFromClientCredentials(input.Resource);
            }

            throw new InvalidOperationException("Unable to authorize without Principal ID or ID Token.");
        }

        /// <summary>
        /// Retrieve Easy Auth token based on provider & principal ID
        /// </summary>
        /// <param name="attribute">The metadata for the token to grab</param>
        /// <returns>Task with Token Store entry of the token</returns>
        private async Task<string> GetEasyAuthAccessTokenAsync(TokenAttribute attribute)
        {
            var jwt = CreateTokenForEasyAuthAccess(attribute);
            EasyAuthTokenStoreEntry tokenStoreEntry = await _easyAuthClient.GetTokenStoreEntryAsync(jwt, attribute);

            bool isTokenValid = IsTokenValid(tokenStoreEntry.AccessToken);
            bool isTokenExpired = tokenStoreEntry.ExpiresOn <= DateTime.UtcNow.AddMinutes(GraphTokenBufferInMinutes);
            bool isRefreshable = IsRefreshableProvider(attribute.IdentityProvider);

            if (isRefreshable && (isTokenExpired || !isTokenValid))
            {
                await _easyAuthClient.RefreshTokenAsync(jwt, attribute);

                // Now that the refresh has occured, grab the new token
                tokenStoreEntry = await _easyAuthClient.GetTokenStoreEntryAsync(jwt, attribute);
            }

            return tokenStoreEntry.AccessToken;
        }

        private async Task<string> GetAuthTokenFromUserTokenAsync(string userToken, string resource)
        {
            if (string.IsNullOrWhiteSpace(resource))
            {
                throw new ArgumentException("A resource is required to get an auth token on behalf of a user.");
            }

            // If the incoming token already has the correct audience (resource), then skip the exchange (it will fail with AADSTS50013!)
            var jwt = new JwtSecurityToken(userToken);
            var currentAudience = GetAudience(jwt);
            if (currentAudience != resource)
            {
                string token = await _aadClient.GetTokenOnBehalfOfUserAsync(
                    userToken,
                    resource);
                return token;
            }

            // No exchange requested, return token directly.
            return userToken;
        }

        private static bool IsTokenValid(string token)
        {
            return JwtHandler.CanReadToken(token);
        }

        private static bool IsRefreshableProvider(string provider)
        {
            //TODO: For now, since we are focusing on AAD, only include it in the refresh path.
            return provider.Equals("AAD", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetAudience(JwtSecurityToken jwt)
        {
            var audience = jwt.Audiences.FirstOrDefault();
            return audience;
        }

        private JwtSecurityToken CreateTokenForEasyAuthAccess(TokenAttribute attribute)
        {
            if (string.IsNullOrEmpty(attribute.UserId))
            {
                throw new ArgumentException("A userId is required to obtain an access token.");
            }

            if (string.IsNullOrEmpty(attribute.IdentityProvider))
            {
                throw new ArgumentException("A provider is necessary to obtain an access token.");
            }

            var identityClaims = new ClaimsIdentity(attribute.UserId);
            identityClaims.AddClaim(new Claim(ClaimTypes.NameIdentifier, attribute.UserId));
            identityClaims.AddClaim(new Claim("idp", attribute.IdentityProvider));

            var hostName = $"https://{_appSettings.Resolve(Constants.AppSettingWebsiteHostname)}/";
            var descr = new SecurityTokenDescriptor
            {
                Audience = hostName,
                Issuer = hostName,
                Subject = identityClaims,
                Expires = DateTime.UtcNow.AddMinutes(_jwtTokenDurationInMinutes),
                SigningCredentials = new HmacSigningCredentials(SigningKey),
            };

            return (JwtSecurityToken)JwtHandler.CreateToken(descr);
        }

    }
}
