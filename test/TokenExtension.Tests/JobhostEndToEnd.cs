// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Threading.Tasks;
    using Microsoft.IdentityModel.Tokens;
    using Moq;
    using TokenBinding;
    using Xunit;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Claims;
    using Microsoft.Azure.WebJobs.Host;
    using System.Reflection;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    public class JobhostEndToEnd
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
            var config = new TokenExtensionConfig();
            var clientFactory = new FakeEasyAuthClientFactory();
            clientFactory.AddResponseInSequence(currentToken);
            config.EasyAuthClientFactory = clientFactory;

            var args = new Dictionary<string, object>
            {
                {"token", "dummyValue" },
            };
            var methodInfo = typeof(TokenFunctions).GetMethods().Where(info => info.Name == "FromId").First();
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);

            await host.CallAsync(methodInfo, args);

            var expectedResult = currentToken.access_token;
            Assert.Equal(expectedResult, finalTokenValue);
            ResetState();
        }

        [Fact]
        public static async Task FromId_TokenExpired_GetRefreshedToken()
        {
            var expiredToken = BuildTokenEntry(DateTime.UtcNow.AddSeconds(-60));
            var refreshedToken = BuildTokenEntry(DateTime.UtcNow.AddDays(1));

            var config = new TokenExtensionConfig();
            var clientFactory = new FakeEasyAuthClientFactory();
            clientFactory.AddResponseInSequence(expiredToken);
            clientFactory.AddResponseInSequence(refreshedToken);
            config.EasyAuthClientFactory = clientFactory;

            var args = new Dictionary<string, object>
            {
                {"token", "dummyValue" },
            };
            var methodInfo = GetMethodInfo("FromId");
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);

            await host.CallAsync(methodInfo, args);

            var expectedResult = refreshedToken.access_token;
            Assert.Equal(expectedResult, finalTokenValue);
            clientFactory.GetMock().Verify(client => client.RefreshToken(It.IsAny<TokenAttribute>()), Times.AtLeastOnce());
            ResetState();
        }

        [Fact]
        public static async Task FromUserToken_CredentialsValid_GetToken()
        {
            var config = new TokenExtensionConfig();
            var clientFactory = new FakeAadClientFactory();
            config.AadClientFactory = clientFactory;

            var nameResolver = new Mock<INameResolver>();
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientIdName)).Returns("dummy");
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientSecretName)).Returns("value");
            config.AppSettings = nameResolver.Object;

            var args = new Dictionary<string, object>
            {
                {"token", SampleUserToken },
            };
            var methodInfo = GetMethodInfo("FromUserToken");
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);
            await host.CallAsync(methodInfo, args);

            var expectedResult = AccessTokenFromUserToken;
            Assert.Equal(expectedResult, finalTokenValue);
            clientFactory.GetMock().Verify(client => client.GetTokenOnBehalfOfUserAsync(SampleUserToken, GraphResource), Times.Exactly(1));
            ResetState();
        }


        [Fact]
        public static async Task UsesAad_MissingClientIdAppSetting_ThrowException()
        {
            var config = new TokenExtensionConfig();
            var clientFactory = new FakeAadClientFactory();
            config.AadClientFactory = clientFactory;

            var nameResolver = new Mock<INameResolver>();
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientIdName)).Returns<string>(null);
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientSecretName)).Returns("value");
            config.AppSettings = nameResolver.Object;

            var args = new Dictionary<string, object>
            {
                {"token", SampleUserToken },
            };
            var methodInfo = GetMethodInfo("FromUserToken");
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);

            await Assert.ThrowsAnyAsync<Exception>(() => host.CallAsync(methodInfo, args));
            ResetState();
        }

        [Fact]
        public static async Task UsesAad_MissingClientSecretAppSetting_ThrowException()
        {
            var config = new TokenExtensionConfig();
            var clientFactory = new FakeAadClientFactory();
            config.AadClientFactory = clientFactory;

            var nameResolver = new Mock<INameResolver>();
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientIdName)).Returns("dummy");
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientSecretName)).Returns<string>(null);
            config.AppSettings = nameResolver.Object;

            var args = new Dictionary<string, object>
            {
                {"token", SampleUserToken },
            };
            var methodInfo = GetMethodInfo("FromUserToken");
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);

            await Assert.ThrowsAnyAsync<Exception>(() => host.CallAsync(methodInfo, args));
            ResetState();
        }

        [Fact]
        public static async Task FromClientCredentials_CredentialsValid_GetToken()
        {
            var config = new TokenExtensionConfig();
            var clientFactory = new FakeAadClientFactory();
            config.AadClientFactory = clientFactory;

            var nameResolver = new Mock<INameResolver>();
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientIdName)).Returns("dummy");
            nameResolver.Setup(x => x.Resolve(Constants.AppSettingClientSecretName)).Returns("value");
            config.AppSettings = nameResolver.Object;

            var args = new Dictionary<string, object>
            {
                { "token", SampleUserToken },
            };
            var methodInfo = GetMethodInfo("ClientCredentials");
            JobHost host = TestHelpers.NewHost<TokenFunctions>(config);
            await host.CallAsync(methodInfo, args);

            var expectedResult = AccessTokenFromClientCredentials;
            Assert.Equal(expectedResult, finalTokenValue);
            clientFactory.GetMock().Verify(client => client.GetTokenFromClientCredentials(GraphResource), Times.Exactly(1));
            ResetState();
        }

        private static void ResetState()
        {
            finalTokenValue = null;
        }

        private static MethodInfo GetMethodInfo(string name)
        {
            return typeof(TokenFunctions).GetMethods().Where(info => info.Name == name).First();
        }

        private static EasyAuthTokenClient.EasyAuthTokenStoreEntry BuildTokenEntry(DateTime expiration)
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
            return new EasyAuthTokenClient.EasyAuthTokenStoreEntry()
            {
                access_token = accessToken,
                expires_on = expiration,
            };
        }

        private class FakeEasyAuthClientFactory : EasyAuthClientFactory
        {
            Mock<IEasyAuthClient> _clientMock;
            Queue<EasyAuthTokenClient.EasyAuthTokenStoreEntry> _responses;

            public FakeEasyAuthClientFactory()
            {
                _clientMock = new Mock<IEasyAuthClient>();
                _responses = new Queue<EasyAuthTokenClient.EasyAuthTokenStoreEntry>();
            }

            internal void AddResponseInSequence(EasyAuthTokenClient.EasyAuthTokenStoreEntry tokenEntry)
            {
                _responses.Enqueue(tokenEntry);
            }

            internal Mock<IEasyAuthClient> GetMock()
            {
                return _clientMock;
            }

            public override IEasyAuthClient GetClient(string hostName, string signingKey, TraceWriter log)
            {
                _clientMock.Setup(client => client.GetTokenStoreEntry(It.IsAny<TokenAttribute>()))
                        .Returns(Task.FromResult(_responses.Dequeue()));
                return _clientMock.Object;
            }
        }

        private class FakeAadClientFactory : AadClientFactory
        {
            Mock<IAadClient> _clientMock;

            public FakeAadClientFactory()
            {
                _clientMock = new Mock<IAadClient>();
                _clientMock.Setup(client => client.GetTokenFromClientCredentials(It.IsAny<string>()))
                    .Returns(Task.FromResult(AccessTokenFromClientCredentials));
                _clientMock.Setup(client => client.GetTokenOnBehalfOfUserAsync(It.IsAny<string>(), It.IsAny<string>()))
                    .Returns(Task.FromResult(AccessTokenFromUserToken));
            }

            public override IAadClient GetClient(ClientCredential credentials)
            {
                return _clientMock.Object;
            }

            public Mock<IAadClient> GetMock()
            {
                return _clientMock;
            }
        }

        private class TokenFunctions
        {
            public void FromId(
                [Token(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = IdentityMode.UserFromId,
                Resource = GraphResource)] string token)
            {
                finalTokenValue = token;
            }

            public void FromUserToken(
                [Token(
                Identity = IdentityMode.UserFromToken,
                UserToken = SampleUserToken,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token)
            {
                finalTokenValue = token;
            }

            public void ClientCredentials(
                [Token(
                Identity = IdentityMode.ClientCredentials,
                UserToken = SampleUserToken,
                IdentityProvider = "AAD",
                Resource = GraphResource)] string token)
            {
                finalTokenValue = token;
            }
        }

    }
}
