// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Net;
    using System.Net.Http;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Azure.WebJobs.Host;
    using Newtonsoft.Json;

    /// <summary>
    /// The client responsible for handling all EasyAuth token-related tasks.
    /// </summary>
    internal class EasyAuthTokenClient : IEasyAuthClient
    {

        internal static readonly JwtSecurityTokenHandler JwtHandler = new JwtSecurityTokenHandler();

        private static readonly HttpClient _httpClient = new HttpClient();

        private static readonly int _jwtExpirationBufferInMinutes = 2;

        private static readonly int _jwtTokenDurationInMinutes = 15;

        private readonly string _baseUrl;

        private readonly string _signingKey;

        private readonly TraceWriter _log;

        private JwtSecurityToken _tokenForEasyAuthAccess;

        /// <summary>
        /// Initializes a new instance of the <see cref="EasyAuthTokenClient"/> class.
        /// </summary>
        /// <param name="hostName">The hostname of the webapp </param>
        /// <param name="signingKey">The website authorization signing key</param>
        public EasyAuthTokenClient(string hostName, string signingKey, TraceWriter log)
        {
            _baseUrl = "https://" + hostName + "/";
            _signingKey = signingKey;
            _log = log;
        }

        private JwtSecurityToken GetTokenForEasyAuthAccess(TokenAttribute attribute)
        {
            if (_tokenForEasyAuthAccess == null || _tokenForEasyAuthAccess.ValidTo <= DateTime.UtcNow.AddMinutes(_jwtExpirationBufferInMinutes))
            {
                _tokenForEasyAuthAccess = CreateTokenForEasyAuthAccess(attribute);
            }

            return _tokenForEasyAuthAccess;
        }

        public async Task<EasyAuthTokenStoreEntry> GetTokenStoreEntry(TokenAttribute attribute)
        {
            var jwt = GetTokenForEasyAuthAccess(attribute);

            // Send the token to the local /.auth/me endpoint and return the JSON
            string meUrl = _baseUrl + $".auth/me?provider={attribute.IdentityProvider}";

            using (var request = new HttpRequestMessage(HttpMethod.Get, meUrl))
            {
                request.Headers.Add("x-zumo-auth", jwt.RawData);
                _log.Verbose($"Fetching user token data from ${meUrl}");
                using (HttpResponseMessage response = await _httpClient.SendAsync(request))
                {
                    _log.Verbose($"Response from '${meUrl}: {response.StatusCode}");
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorResponse = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException($"Request to {_baseUrl} failed. Status Code: {response.StatusCode}; Body: {errorResponse}");
                    }
                    var responseString = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<EasyAuthTokenStoreEntry>(responseString);
                }
            }
        }

        public async Task RefreshToken(TokenAttribute attribute)
        {
            if (string.IsNullOrEmpty(attribute.Resource))
            {
                throw new ArgumentException("A resource is required to renew an access token.");
            }

            if (string.IsNullOrEmpty(attribute.UserId))
            {
                throw new ArgumentException("A userId is required to renew an access token.");
            }

            if (string.IsNullOrEmpty(attribute.IdentityProvider))
            {
                throw new ArgumentException("A provider is necessary to renew an access token.");
            }

            string refreshUrl = _baseUrl + $".auth/refresh?resource=" + WebUtility.UrlEncode(attribute.Resource);

            using (var refreshRequest = new HttpRequestMessage(HttpMethod.Get, refreshUrl))
            {
                var jwt = GetTokenForEasyAuthAccess(attribute);
                refreshRequest.Headers.Add("x-zumo-auth", jwt.RawData);
                _log.Verbose($"Refreshing ${attribute.IdentityProvider} access token for user ${attribute.UserId} at ${refreshUrl}");
                using (HttpResponseMessage response = await _httpClient.SendAsync(refreshRequest))
                {
                    _log.Verbose($"Response from ${refreshUrl}: {response.StatusCode}");
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorResponse = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException($"Failed to refresh {attribute.UserId} {attribute.IdentityProvider} error={response.StatusCode} {errorResponse}");
                    }
                }
            }
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

            var descr = new SecurityTokenDescriptor
            {
                Audience = _baseUrl,
                Issuer = _baseUrl,
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

            protected static string CreateDigestAlgorithm(byte[] key)
            {
                if (key.Length <= 32)
                {
                    return Algorithms.Sha256Digest;
                }
                else if (key.Length <= 48)
                {
                    return Algorithms.Sha384Digest;
                }
                else
                {
                    return Algorithms.Sha512Digest;
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

                public const string Sha256Digest = "http://www.w3.org/2001/04/xmlenc#sha256";
                public const string Sha384Digest = "http://www.w3.org/2001/04/xmlenc#sha384";
                public const string Sha512Digest = "http://www.w3.org/2001/04/xmlenc#sha512";
            }
        }
    }
}
