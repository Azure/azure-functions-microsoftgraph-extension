// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.IO;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Options;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
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

        private static string AccessTokenFromClientCredentials = "clientcredentials";

        private static string AccessTokenFromUserToken = "usertoken";

        private static string finalTokenValue;       

        [Fact]
        public static async Task FromId_TokenStillValid_GetStoredToken()
        {
            
            var currentToken = BuildTokenEntry(DateTime.UtcNow.AddDays(1));
            IEasyAuthClient mockEasyAuthClient = GetEasyAuthClientMock(currentToken).Object;
            IAadClient aadClient = GetAadClientMock().Object;
            INameResolver appSettings = GetNameResolver(new Dictionary<string, string>()
            {
                { Constants.WebsiteAuthSigningKey, SigningKey }
            }).Object;

            OutputContainer outputContainer = await TestHelpers.RunTestAsync<TokenFunctions>("TokenFunctions.FromId",  appSettings: appSettings, easyAuthClient: mockEasyAuthClient, aadClient: aadClient);

            var expectedResult = currentToken.AccessToken;
            Assert.Equal(expectedResult, outputContainer.Output);
            ResetState();
        }

        [Fact]
        public static async Task FromId_TokenExpired_GetRefreshedToken()
        {
            var expiredToken = BuildTokenEntry(DateTime.UtcNow.AddSeconds(-60));
            var refreshedToken = BuildTokenEntry(DateTime.UtcNow.AddDays(1));

            var mockClient = GetEasyAuthClientMock(expiredToken, refreshedToken);
            INameResolver appSettings = GetNameResolver(new Dictionary<string, string>()
            {
                { Constants.WebsiteAuthSigningKey, SigningKey }
            }).Object;

            OutputContainer outputContainer = await TestHelpers.RunTestAsync<TokenFunctions>("TokenFunctions.FromId", appSettings: appSettings, easyAuthClient: mockClient.Object);

            var expectedResult = refreshedToken.AccessToken;
            Assert.Equal(expectedResult, outputContainer.Output);
            mockClient.Verify(client => client.RefreshToken(It.IsAny<JwtSecurityToken>(), It.IsAny<TokenAttribute>()), Times.AtLeastOnce());
            ResetState();
        }

        [Fact]
        public static async Task FromUserToken_CredentialsValid_GetToken()
        {
            var mockClient = GetAadClientMock();

            OutputContainer outputContainer = await TestHelpers.RunTestAsync<TokenFunctions>("TokenFunctions.FromUserToken", aadClient: mockClient.Object);

            var expectedResult = AccessTokenFromUserToken;
            Assert.Equal(expectedResult, outputContainer.Output);
            mockClient.Verify(client => client.GetTokenOnBehalfOfUserAsync(SampleUserToken, GraphResource), Times.Exactly(1));
            ResetState();
        }

        [Fact]
        public static async Task FromClientCredentials_CredentialsValid_GetToken()
        {
            var mockClient = GetAadClientMock();

            var args = new Dictionary<string, object>
            {
                { "token", SampleUserToken },
            };

            OutputContainer outputContainer = await TestHelpers.RunTestAsync<TokenFunctions>("TokenFunctions.ClientCredentials", aadClient: mockClient.Object);

            var expectedResult = AccessTokenFromClientCredentials;
            Assert.Equal(expectedResult, outputContainer.Output);
            mockClient.Verify(client => client.GetTokenFromClientCredentials(GraphResource), Times.Exactly(1));
            ResetState();
        }

        [Fact]
        public static async Task Integrated_FromClientCredentials_CredentialsValid_GetToken()
        {
            var options = TestHelpers.GetValidSettingsForTests();
            IAadClient aadClient = new AadClient(Options.Create(options));

            OutputContainer outputContainer = await TestHelpers.RunTestAsync<RealTokenFunctions>("RealTokenFunctions.ClientCredentials", aadClient: aadClient);

            var token = new JwtSecurityToken((string) outputContainer.Output);

            Assert.True(token.ValidTo > DateTime.UtcNow);
            Assert.True(token.Audiences.Contains(GraphResource));
        }

        [Fact]
        public static async Task Integrated_FromClientCredentials_ClientSecretInvalid_GetToken()
        {
            var options = TestHelpers.GetValidSettingsForTests();
            options.ClientSecret = "invalid";

            IAadClient aadClient = new AadClient(Options.Create(options));

            try
            {
                OutputContainer outputContainer = await TestHelpers.RunTestAsync<RealTokenFunctions>("RealTokenFunctions.ClientCredentials", aadClient: aadClient);              
            }
            catch(Host.FunctionInvocationException e)
            {
                Assert.True(e.InnerException.InnerException is AdalServiceException);
                Assert.True(e.InnerException.InnerException.Message.Contains("Invalid client secret"));
            }
        }

        [Fact]
        public static async Task Integrated_FromClientCredentials_ClientIDInvalid_GetToken()
        {
            var options = TestHelpers.GetValidSettingsForTests();
            options.ClientId = "invalid";

            IAadClient aadClient = new AadClient(Options.Create(options));

            try
            {
                OutputContainer outputContainer = await TestHelpers.RunTestAsync<RealTokenFunctions>("RealTokenFunctions.ClientCredentials", aadClient: aadClient);
            }
            catch (Host.FunctionInvocationException e)
            {
                Assert.True(e.InnerException.InnerException is AdalServiceException);
                Assert.True(e.InnerException.InnerException.Message.Contains("not found"));
            }
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
                SigningCredentials = new EasyAuthTokenManager.HmacSigningCredentials(SigningKey),
            };
            string accessToken = (new JwtSecurityTokenHandler()).CreateJwtSecurityToken(descr).RawData;
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
            clientMock.Setup(client => client.GetTokenStoreEntry(It.IsAny<JwtSecurityToken>(), It.IsAny<TokenAttribute>()))
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

        private static Mock<INameResolver> GetNameResolver(Dictionary<string, string> appSettings)
        {
            var mock = new Mock<INameResolver>();
            foreach(var appSetting in appSettings)
            {
                mock.Setup(resolver => resolver.Resolve(appSetting.Key)).Returns(appSetting.Value);
            }
            return mock;
        }

        public class TokenFunctions
        {
            public void FromId(
                [Token(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Resource = GraphResource)] string token, OutputContainer outputContainer)
            {
                outputContainer.Output = token;
            }

            public void FromUserToken(
                [Token(
                Identity = TokenIdentityMode.UserFromToken,
                UserToken = SampleUserToken,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token, OutputContainer outputContainer)
            {
                outputContainer.Output = token;
            }

            public void ClientCredentials(
                [Token(
                Identity = TokenIdentityMode.ClientCredentials,
                UserToken = SampleUserToken,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token, OutputContainer outputContainer)
            {
                outputContainer.Output = token;
            }
        }

        public class RealTokenFunctions
        {
            public void ClientCredentials(
                [Token(
                Identity = TokenIdentityMode.ClientCredentials,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token, OutputContainer outputContainer)
            {
                outputContainer.Output = token;
            }
        }
    }
}