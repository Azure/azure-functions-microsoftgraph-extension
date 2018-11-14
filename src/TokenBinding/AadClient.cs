// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.Services.AppAuthentication;
    using Microsoft.Extensions.Options;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    /// <summary>
    /// Helper class for calling onto ADAL
    /// </summary>
    internal class AadClient : IAadClient
    {
        private AuthenticationContext _authContext;
        private ClientCredential _clientCredentials;
        private TokenOptions _options;

        private AuthenticationContext AuthContext {
            get
            {
                if (_authContext == null)
                {
                    // NOTE: We had to turn off authority validation here, otherwise we would
                    // get the error "AADSTS50049: Unknown or invalid instance" for some tenants.
                    _authContext = new AuthenticationContext(_options.TenantUrl, false);
                }
                return _authContext;
            }
        }

        private ClientCredential ClientCredentials
        {
            get
            {
                if (_clientCredentials == null)
                {
                    if (_options.ClientId == null || _options.ClientSecret == null)
                    {
                        throw new InvalidOperationException($"The Application Settings {Constants.ClientIdName} and {Constants.ClientSecretName} must be set to perform this operation.");
                    }
                    _clientCredentials = new ClientCredential(_options.ClientId, _options.ClientSecret);
                }

                return _clientCredentials;
            }
        }

        public AadClient(IOptions<TokenOptions> options)
        {
            _options = options.Value;
        }

        /// <summary>
        /// Use client credentials to retrieve auth token
        /// Typically used to retrieve a token for a different audience
        /// </summary>
        /// <param name="userToken">User's token for a given resource</param>
        /// <param name="resource">Resource the token is for (e.g. https://graph.microsoft.com)</param>
        /// <returns>Access token for correct audience</returns>
        public async Task<string> GetTokenOnBehalfOfUserAsync(
            string userToken,
            string resource)
        {
            if (string.IsNullOrEmpty(userToken))
            {
                throw new ArgumentException("A usertoken is required to retrieve a token for a user.");
            }

            if (string.IsNullOrEmpty(resource))
            {
                throw new ArgumentException("A resource is required to retrieve a token for a user.");
            }

            UserAssertion userAssertion = new UserAssertion(userToken);
            AuthenticationResult ar = await AuthContext.AcquireTokenAsync(resource, ClientCredentials, userAssertion);
            return ar.AccessToken;
        }

        public async Task<string> GetTokenFromClientCredentials(string resource)
        {
            if (string.IsNullOrEmpty(resource))
            {
                throw new ArgumentException("A resource is required to retrieve a token from client credentials.");
            }

            AuthenticationResult authResult = await AuthContext.AcquireTokenAsync(resource, ClientCredentials);
            return authResult.AccessToken;
        }

        public async Task<string> GetTokenFromAppIdentity(string resource)
        {
            if (string.IsNullOrEmpty(resource))
            {
                throw new ArgumentException("A resource is required to retrieve a token from the application's identity via MSI.");
            }

            var azureServiceTokenProvider = new AzureServiceTokenProvider();

            var token = await azureServiceTokenProvider.GetAccessTokenAsync(resource);

            return token;
        }
    }
}
