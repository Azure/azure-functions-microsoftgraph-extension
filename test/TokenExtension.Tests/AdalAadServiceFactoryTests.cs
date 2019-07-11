using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    public class AdalAadServiceFactoryTests
    {
        string validUrl = "https://login.microsoft.net/common";
        string clientId = "id";
        string clientSecret = "secret";

        [Fact]
        public void CreateAadService_MissingTenantUrl_ThrowsException()
        {
            AdalAadServiceFactory adalAadServiceFactory = new AdalAadServiceFactory();
            // Assert throws when null
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(null, clientId, clientSecret));
            // Assert throws when empty string
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(string.Empty, clientId, clientSecret));
            // Assert throws when whitespace
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(" ", clientId, clientSecret));
        }

        [Fact]
        public void CreateAadService_InvalidTenantUrl_ThrowsException()
        {
            AdalAadServiceFactory adalAadServiceFactory = new AdalAadServiceFactory();
            // Assert throws when not a url
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient("noturl", clientId, clientSecret));
            // Assert throws when not at least 1 path segment
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient("https://login.microsoft.net", clientId, clientSecret));
            // Assert throws when not https
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient("http://login.microsoft.net/common", clientId, clientSecret));
        }

        [Fact]
        public void CreateAadService_MissingClientId_ThrowsException()
        {
            AdalAadServiceFactory adalAadServiceFactory = new AdalAadServiceFactory();
            // Assert throws when null
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(validUrl, null, clientSecret));
            // Assert throws when empty string
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(validUrl, string.Empty, clientSecret));
            // Assert throws when whitespace
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(validUrl, " ", clientSecret));
        }

        [Fact]
        public void CreateAadService_MissingClientSecret_ThrowsException()
        {
            AdalAadServiceFactory adalAadServiceFactory = new AdalAadServiceFactory();
            // Assert throws when null
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(validUrl, clientId, null));
            // Assert throws when empty string
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(validUrl, clientId, string.Empty));
            // Assert throws when whitespace
            Assert.ThrowsAny<Exception>(() => adalAadServiceFactory.GetAadClient(validUrl, clientId, " "));
        }

        [Fact]
        public void CreateAadService_AllRequiredParameters_ReturnsAdalAadService()
        {
            AdalAadServiceFactory adalAadServiceFactory = new AdalAadServiceFactory();
            AdalAadService aadService = adalAadServiceFactory.GetAadClient("https://login.microsoft.net/common", clientId, clientSecret) as AdalAadService;
            Assert.NotNull(aadService);
        }
    }
}
