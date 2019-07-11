using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
using Microsoft.IdentityModel.Tokens;
using Moq;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    internal static class MockHelper
    {

        //Dummy Key
        public const string SigningKey = @"MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCK91gD4uDPt
        a68CT8Mzhcoqv3zaA0hnrJLXQMdddXKbmvdRRhQzujnF3NwF6M5BIbX9o8+Z6GUMTH14l1/hk3z8zA5aNCqjO
        QOgkjZUZwBynD6KgLPBn+ilZqpQDYEEnzZo34Y99zOtorS+GyJTcQeo4jiOzTVLR6I7GrMEuHqnwIDAQAB";


        public const string GraphResource = "https://graph.microsoft.com";

        public const string SampleUserToken = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiYWRtaW4iOnRydWV9.TJVA95OrM7E2cBab30RMHrHDcEfxjoYZgeFONFh7HgQ";

        public static string AccessTokenFromClientCredentials = "clientcredentials";

        public static string AccessTokenFromUserToken = "usertoken";

        public static EasyAuthTokenStoreEntry BuildTokenEntry(DateTime expiration)
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

        public static Mock<IEasyAuthClient> GetEasyAuthClientMock(params EasyAuthTokenStoreEntry[] responsesInOrder)
        {
            var clientMock = new Mock<IEasyAuthClient>();
            var responseQueue = new Queue<EasyAuthTokenStoreEntry>(responsesInOrder);
            clientMock.Setup(client => client.GetTokenStoreEntry(It.IsAny<JwtSecurityToken>(), It.IsAny<TokenAttribute>()))
                .Returns(Task.FromResult(responseQueue.Dequeue()));
            return clientMock;
        }

        public static Mock<IAadService> GetAadClientMock()
        {
            var clientMock = new Mock<IAadService>();
            clientMock.Setup(client => client.GetTokenFromClientCredentials(It.IsAny<string>()))
                .Returns(Task.FromResult(AccessTokenFromClientCredentials));
            clientMock.Setup(client => client.GetTokenOnBehalfOfUserAsync(It.IsAny<string>(), It.IsAny<string>()))
                .Returns(Task.FromResult(AccessTokenFromUserToken));
            return clientMock;
        }

        public static IAadServiceFactory GetAadServiceFactory(Mock<IAadService> serviceMock)
        {
            IAadService service = serviceMock.Object;
            var factoryMock = new Mock<IAadServiceFactory>();
            factoryMock.Setup(factory => factory.GetAadClient(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Returns(service);
            return factoryMock.Object;
        }

        public static Mock<INameResolver> GetAppSettingsMock(IDictionary<string, string> appSettings)
        {
            var appSettingMock = new Mock<INameResolver>();
            foreach(var appSetting in appSettings)
            {
                appSettingMock.Setup(resolver => resolver.Resolve(appSetting.Key)).Returns(appSetting.Value);
            }
            return appSettingMock;
        } 
    }
}
