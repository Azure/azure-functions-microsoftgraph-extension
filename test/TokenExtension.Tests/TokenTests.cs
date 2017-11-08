// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.IdentityModel.Tokens;
    using Moq;
    using Xunit;

    public class TokenTests
    {
        //Dummy Key
        private const string SigningKey = @"MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCK91gD4uDPt
        a68CT8Mzhcoqv3zaA0hnrJLXQMdddXKbmvdRRhQzujnF3NwF6M5BIbX9o8+Z6GUMTH14l1/hk3z8zA5aNCqjO
        QOgkjZUZwBynD6KgLPBn+ilZqpQDYEEnzZo34Y99zOtorS+GyJTcQeo4jiOzTVLR6I7GrMEuHqnwIDAQAB";

        private const string GraphResource = "https://graph.microsoft.com";

        private const string SampleUserToken = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiYWRtaW4iOnRydWV9.TJVA95OrM7E2cBab30RMHrHDcEfxjoYZgeFONFh7HgQ";

        private static readonly JwtSecurityTokenHandler jwtHandler = new JwtSecurityTokenHandler();

        private static string AccessTokenFromClientCredentials = "clientcredentials";

        private static string AccessTokenFromUserToken = "usertoken";

        private static string finalTokenValue;

        [Fact]
        public static async Task FromId_TokenStillValid_GetStoredToken()
        {
            var currentToken = BuildTokenEntry(DateTime.UtcNow.AddDays(1));
            var config = new AuthTokenExtensionConfig();
            var mockClient = GetEasyAuthClientMock(currentToken);
            config.EasyAuthClient = mockClient.Object;

            var args = new Dictionary<string, object>
            {
                {"token", "dummyValue" },
            };
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);
            await host.CallAsync("TokenFunctions.FromId", args);

            var expectedResult = currentToken.AccessToken;
            Assert.Equal(expectedResult, finalTokenValue);
            ResetState();
        }

        [Fact]
        public static async Task FromId_TokenExpired_GetRefreshedToken()
        {
            var expiredToken = BuildTokenEntry(DateTime.UtcNow.AddSeconds(-60));
            var refreshedToken = BuildTokenEntry(DateTime.UtcNow.AddDays(1));

            var config = new AuthTokenExtensionConfig();
            var mockClient = GetEasyAuthClientMock(expiredToken, refreshedToken);
            config.EasyAuthClient = mockClient.Object;

            var args = new Dictionary<string, object>
            {
                {"token", "dummyValue" },
            };
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);
            await host.CallAsync("TokenFunctions.FromId", args);

            var expectedResult = refreshedToken.AccessToken;
            Assert.Equal(expectedResult, finalTokenValue);
            mockClient.Verify(client => client.RefreshToken(It.IsAny<TokenAttribute>()), Times.AtLeastOnce());
            ResetState();
        }

        [Fact]
        public static async Task FromUserToken_CredentialsValid_GetToken()
        {
            var config = new AuthTokenExtensionConfig();
            var mockClient = GetAadClientMock();
            config.AadClient = mockClient.Object;

            var args = new Dictionary<string, object>
            {
                {"token", SampleUserToken },
            };
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);
            await host.CallAsync("TokenFunctions.FromUserToken", args);

            var expectedResult = AccessTokenFromUserToken;
            Assert.Equal(expectedResult, finalTokenValue);
            mockClient.Verify(client => client.GetTokenOnBehalfOfUserAsync(SampleUserToken, GraphResource), Times.Exactly(1));
            ResetState();
        }

        [Fact]
        public static async Task FromClientCredentials_CredentialsValid_GetToken()
        {
            var config = new AuthTokenExtensionConfig();
            var mockClient = GetAadClientMock();
            config.AadClient = mockClient.Object;

            var args = new Dictionary<string, object>
            {
                { "token", SampleUserToken },
            };
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);
            await host.CallAsync("TokenFunctions.ClientCredentials", args);

            var expectedResult = AccessTokenFromClientCredentials;
            Assert.Equal(expectedResult, finalTokenValue);
            mockClient.Verify(client => client.GetTokenFromClientCredentials(GraphResource), Times.Exactly(1));
            ResetState();
        }

        private static void ResetState()
        {
            finalTokenValue = null;
        }

        private static EasyAuthTokenStoreEntry BuildTokenEntry(DateTime expiration)
        {
            ClaimsIdentity identity = new ClaimsIdentity();
            identity.AddClaim(new Claim(ClaimTypes.Name, "Sample"));
            identity.AddClaim(new Claim("idp", "aad"));

            var descr = new SecurityTokenDescriptor
            {
                Audience = "https://sample.com",
                Issuer = "https://sample.com",
                Subject = identity,
                SigningCredentials = new EasyAuthTokenClient.HmacSigningCredentials(SigningKey),
            };
            string accessToken = jwtHandler.CreateJwtSecurityToken(descr).RawEncryptedKey;
            return new EasyAuthTokenStoreEntry()
            {
                AccessToken = accessToken,
                ExpiresOn = expiration,
            };
        }

        private static Mock<IEasyAuthClient> GetEasyAuthClientMock(params EasyAuthTokenStoreEntry[] responsesInOrder)
        {
            var clientMock = new Mock<IEasyAuthClient>();
            var responseQueue = new Queue<EasyAuthTokenStoreEntry>(responsesInOrder);
            clientMock.Setup(client => client.GetTokenStoreEntry(It.IsAny<TokenBaseAttribute>()))
                .Returns(Task.FromResult(responseQueue.Dequeue()));
            return clientMock;
        }

        private static Mock<IAadClient> GetAadClientMock()
        {
            var clientMock = new Mock<IAadClient>();
            clientMock.Setup(client => client.GetTokenFromClientCredentials(It.IsAny<string>()))
                .Returns(Task.FromResult(AccessTokenFromClientCredentials));
            clientMock.Setup(client => client.GetTokenOnBehalfOfUserAsync(It.IsAny<string>(), It.IsAny<string>()))
                .Returns(Task.FromResult(AccessTokenFromUserToken));
            return clientMock;
        }

        private class TokenFunctions
        {
            public void FromId(
                [Token(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Resource = GraphResource)] string token)
            {
                finalTokenValue = token;
            }

            public void FromUserToken(
                [Token(
                Identity = TokenIdentityMode.UserFromToken,
                UserToken = SampleUserToken,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token)
            {
                finalTokenValue = token;
            }

            public void ClientCredentials(
                [Token(
                Identity = TokenIdentityMode.ClientCredentials,
                UserToken = SampleUserToken,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token)
            {
                finalTokenValue = token;
            }
        }

    }
}