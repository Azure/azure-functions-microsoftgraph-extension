// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Net.Http.Headers;
    using Moq;
    using Xunit;
    using static Microsoft.Azure.WebJobs.Extensions.AuthTokens.EasyAuthTokenManager;

    public class EasyAuthTokenTests
    {

        internal static JwtSecurityTokenHandler JwtHandler = new JwtSecurityTokenHandler();

        [Fact]
        public static async Task FromUserRequest_EasyAuthNotEnabled_Fails()
        {
            var bearerTokenRequest = GetBearerTokenRequest();
            var appSettings = new Dictionary<string, string>()
            {
                { Constants.ClientIdName, Guid.NewGuid().ToString() },
                { Constants.ClientSecretName, "secret" },
                { Constants.WebsiteAuthOpenIdIssuer, "https://login.microsoft.net/common" },
            };
            OutputContainer outputContainer = await TestHelpers.RunTestAsync<TokenFunctions>(nameof(TokenFunctions.FromUserRequest), options: appSettings, request: bearerTokenRequest);

            var expectedResult = MockHelper.AccessTokenFromUserToken;
            Assert.Equal(expectedResult, outputContainer.Output);
        }

        [Fact]
        public static async Task FromUserRequest_ValidBearerToken_GetGraphToken()
        {
            var bearerTokenRequest = GetBearerTokenRequest();
            OutputContainer outputContainer = await TestHelpers.RunTestAsync<TokenFunctions>(nameof(TokenFunctions.FromUserRequest), request: bearerTokenRequest);

            var expectedResult = MockHelper.AccessTokenFromUserToken;
            Assert.Equal(expectedResult, outputContainer.Output);
        }

        [Fact]
        public static async Task FromUserRequest_AccessTokenValidJwt_ClientIdAudience_GetGraphToken()
        {
        }

        [Fact]
        public static async Task FromUserRequest_AccessTokenValidJwt_GraphAudience_ReturnAccessToken()
        {
        }

        [Fact]
        public static async Task FromUserRequest_AccessTokenValidJwt_UnrelatedAudience_ReturnAccessToken()
        {
        }

        [Fact]
        public static async Task FromUserRequest_AccessTokenExpiredJwt_HasRefreshToken_GetValidGraphToken()
        {
        }

        // Not sure if this scenario is technically possible?
        [Fact]
        public static async Task FromUserRequest_AccessTokenExpiredJwt_NoRefreshToken_Fails()
        {
        }


        // Not sure if this scenario is technically possible?
        [Fact]
        public static async Task FromUserRequest_AccessTokenNotJwt_NoRefreshToken_Fails()
        {
        }

        [Fact]
        public static async Task FromUserRequest_AccessTokenNotJwt_HasRefreshToken_GetValidGraphToken()
        {
        }


        private static Mock<INameResolver> GetNameResolver(Dictionary<string, string> appSettings)
        {
            var mock = new Mock<INameResolver>();
            foreach(var appSetting in appSettings)
            {
                mock.Setup(resolver => resolver.Resolve(appSetting.Key)).Returns(appSetting.Value);
            }
            return mock;
        }

        private static string GetJwtToken(string audience, DateTime? expiration = null, string version = "1.0", string oid = null)
        {
            ClaimsIdentity identity = new ClaimsIdentity();
            identity.AddClaim(new Claim("ver", version));
            if (oid != null)
            {
                identity.AddClaim(new Claim("oid", oid));
            }

            var descr = new SecurityTokenDescriptor
            {
                Audience = audience,
                Issuer = "https://login.microsoftonline.com/common",
                Expires = expiration ?? DateTime.UtcNow.AddHours(1),
                SigningCredentials = new HmacSigningCredentials(MockHelper.SigningKey),
                Subject = identity, 
            };

            var jwt = JwtHandler.CreateJwtSecurityToken(descr);
            return jwt.RawData;
        }

        // Simulates the Bearer token authentication flow for EasyAuth with AAD
        private static HttpRequest GetBearerTokenRequest()
        {
            var jwt = GetJwtToken(MockHelper.GraphResource);
            return GetHttpRequest(new Dictionary<string, string>() { { HeaderNames.Authorization, $"Bearer {jwt}" } });
        }

        // Simulate the Server-directed flow or client-directed flow for EasyAuth with AAD
        private static HttpRequest GetTokenStorePopulatedRequest(AccessTokenDetails accessToken, bool populateRefreshToken = true)
        {
            string oid = Guid.NewGuid().ToString();
            string AccessTokenHeaderName = "X-MS-TOKEN-AAD-ACCESS-TOKEN";
            string RefreshTokenHeaderName = "X-MS-TOKEN-AAD-REFRESH-TOKEN";
            string IdTokenHeaderName = "X-MS-TOKEN-AAD-ID-TOKEN";
            string AccessTokenExpirationHeaderName = "X-MS-TOKEN-AAD-EXPIRES-ON";

            IDictionary<string, string> headers = new Dictionary<string, string>();

            if (accessToken.IsJwt)
            {
                headers.Add(AccessTokenHeaderName, GetJwtToken(accessToken.Audience, accessToken.Expiration, accessToken.Version, oid));
            }
            else
            {
                headers.Add(AccessTokenHeaderName, "nonsense"); // We don't know the format of these non JWT tokens, so don't bother trying to emulate it.
            }

            string expirationDateTimeFormat = "yyyy'-'MM'-'dd HH':'mm':'ss'Z'";
            var expirationTimeString = (accessToken.Expiration ?? DateTime.UtcNow.AddHours(1)).ToString(expirationDateTimeFormat);
            headers.Add(AccessTokenExpirationHeaderName, expirationTimeString);

            headers.Add(IdTokenHeaderName, GetJwtToken("client-id", accessToken.Expiration, accessToken.Version, oid));

            if (populateRefreshToken)
            {
                headers.Add(RefreshTokenHeaderName, "nonsense"); // We don't know the format of these non JWT tokens, so don't bother trying to emulate it.
            }

            return GetHttpRequest(headers);
        }

        private static HttpRequest GetHttpRequest(IDictionary<string, string> headers)
        {
            HttpContext context = new DefaultHttpContext();
            HttpRequest req = context.Request;
            foreach(var keyValue in headers)
            {
                req.Headers[keyValue.Key] = keyValue.Value;
            }
            return req;
        }

        private struct AccessTokenDetails
        {
            public bool IsJwt;
            public DateTime? Expiration;
            public string Version;
            public string Audience;
        }

        public class TokenFunctions
        {
            public void FromUserRequest(
                [Token(
                Identity = TokenIdentityMode.UserFromRequest,
                AadResource = MockHelper.GraphResource)] string token, OutputContainer outputContainer)
            {
                outputContainer.Output = token;
            }
        }
    }
}