using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    internal class AdalAadServiceFactory : IAadServiceFactory
    {
        public IAadService GetAadClient(string tenantUrl, string clientId, string clientSecret)
        {
            if (string.IsNullOrWhiteSpace(tenantUrl) || string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(clientSecret))
            {
                throw new InvalidOperationException($"Cannot use {TokenIdentityMode.ClientCredentials} if any of {Constants.WebsiteAuthOpenIdIssuer}, {Constants.ClientIdName}, {Constants.ClientSecretName} are null.");
            }

            AuthenticationContext authenticationContext = new AuthenticationContext(tenantUrl, false);
            ClientCredential credential = new ClientCredential(clientId, clientSecret);
            return new AdalAadService(authenticationContext, credential);
        }
    }
}
