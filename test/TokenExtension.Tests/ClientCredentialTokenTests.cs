using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
using Moq;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    public class ClientCredentialTokenTests
    {
        [Fact]
        public static async Task FromClientCredentials_CredentialsValid_GetGraphToken()
        {
            var mockClient = MockHelper.GetAadClientMock();
            IAadServiceFactory aadServiceFactory = MockHelper.GetAadServiceFactory(mockClient);

            OutputContainer outputContainer = await TestHelpers.RunTestAsync<ClientCredentialFunctions>(nameof(ClientCredentialFunctions.ClientCredentials), aadServiceFactory: aadServiceFactory);

            var expectedResult = MockHelper.AccessTokenFromClientCredentials;
            Assert.Equal(expectedResult, outputContainer.Output);
            mockClient.Verify(client => client.GetTokenFromClientCredentials(MockHelper.GraphResource), Times.Exactly(1));
        }

        public class ClientCredentialFunctions
        {
            public void ClientCredentials(
                [Token(
                Identity = TokenIdentityMode.ClientCredentials,
                AadResource = MockHelper.GraphResource)] string token, OutputContainer outputContainer)
            {
                outputContainer.Output = token;
            }
        }
    }
}
