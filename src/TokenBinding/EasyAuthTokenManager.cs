// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Host;
    using Microsoft.IdentityModel.Tokens;

    /// <summary>
    /// Class representing an application's [EasyAuth] Token Store
    /// see  https://cgillum.tech/2016/03/07/app-service-token-store/
    /// </summary>
    internal class EasyAuthTokenManager
    {
        internal static readonly JwtSecurityTokenHandler JwtHandler = new JwtSecurityTokenHandler();

  
        private static readonly int GraphTokenBufferInMinutes = 5;

        private static readonly int _jwtExpirationBufferInMinutes = 2;

        private static readonly int _jwtTokenDurationInMinutes = 15;

        private readonly string _signingKey;

        private readonly IEasyAuthClient _client;

        /// <summary>
        /// Initializes a new instance of the <see cref="EasyAuthTokenManager"/> class.
        /// </summary>
        /// <param name="hostName">The hostname of the keystore. </param>
        /// <param name="signingKey">The website authorization signing key</param>
        public EasyAuthTokenManager(IEasyAuthClient client, string signingKey)
        {
            _client = client;
            _signingKey = signingKey;
        }

        /// <summary>
        /// Retrieve Easy Auth token based on provider & principal ID
        /// </summary>
        /// <param name="attribute">The metadata for the token to grab</param>
        /// <returns>Task with Token Store entry of the token</returns>
        public async Task<string> GetEasyAuthAccessTokenAsync(TokenAttribute attribute)
        {
            var jwt = CreateTokenForEasyAuthAccess(attribute);
            EasyAuthTokenStoreEntry tokenStoreEntry = await _client.GetTokenStoreEntry(jwt, attribute);

            bool isTokenValid = IsTokenValid(tokenStoreEntry.AccessToken);
            bool isTokenExpired = tokenStoreEntry.ExpiresOn <= DateTime.UtcNow.AddMinutes(GraphTokenBufferInMinutes);
            bool isRefreshable = IsRefreshableProvider(attribute.IdentityProvider);

            if (isRefreshable && (isTokenExpired || !isTokenValid))
            {
                await _client.RefreshToken(jwt, attribute);

                // Now that the refresh has occured, grab the new token
                tokenStoreEntry = await _client.GetTokenStoreEntry(jwt, attribute);
            }

            return tokenStoreEntry.AccessToken;
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
            identityClaims.AddClaim(new System.Security.Claims.Claim(ClaimTypes.NameIdentifier, attribute.UserId));
            identityClaims.AddClaim(new System.Security.Claims.Claim("idp", attribute.IdentityProvider));

            var baseUrl = _client.GetBaseUrl();
            var descr = new SecurityTokenDescriptor
            {
                Audience = baseUrl,
                Issuer = baseUrl,
                Subject = identityClaims,
                Expires = DateTime.UtcNow.AddMinutes(_jwtTokenDurationInMinutes),
                SigningCredentials = new HmacSigningCredentials(_signingKey),
            };

            return (JwtSecurityToken)JwtHandler.CreateToken(descr);
        }


        public class HmacSigningCredentials : SigningCredentials
        {
            public HmacSigningCredentials(string base64EncodedKey)
                : this(ParseKeyString(base64EncodedKey))
            { }

            public HmacSigningCredentials(byte[] key)
                : base(new SymmetricSecurityKey(key), CreateSignatureAlgorithm(key))
            {
            }

            /// <summary>
            /// Converts a base64 OR hex-encoded string into a byte array.
            /// </summary>
            protected static byte[] ParseKeyString(string keyString)
            {
                if (string.IsNullOrEmpty(keyString))
                {
                    return new byte[0];
                }
                else if (IsHexString(keyString))
                {
                    return HexStringToByteArray(keyString);
                }
                else
                {
                    return Convert.FromBase64String(keyString);
                }
            }

            protected static bool IsHexString(string input)
            {
                if (string.IsNullOrEmpty(input))
                {
                    throw new ArgumentNullException(nameof(input));
                }

                for (int i = 0; i < input.Length; i++)
                {
                    char c = input[i];
                    bool isHexChar = (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
                    if (!isHexChar)
                    {
                        return false;
                    }
                }

                return true;
            }

            protected static byte[] HexStringToByteArray(string hexString)
            {
                byte[] bytes = new byte[hexString.Length / 2];
                for (int i = 0; i < hexString.Length; i += 2)
                {
                    bytes[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
                }

                return bytes;
            }

            protected static string CreateSignatureAlgorithm(byte[] key)
            {
                if (key.Length <= 32)
                {
                    return Algorithms.HmacSha256Signature;
                }
                else if (key.Length <= 48)
                {
                    return Algorithms.HmacSha384Signature;
                }
                else
                {
                    return Algorithms.HmacSha512Signature;
                }
            }

            /// <summary>
            /// Key value pairs (algorithm name, w3.org link)
            /// </summary>
            protected static class Algorithms
            {
                public const string HmacSha256Signature = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha256";
                public const string HmacSha384Signature = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha384";
                public const string HmacSha512Signature = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha512";
            }
        }
    }
}
