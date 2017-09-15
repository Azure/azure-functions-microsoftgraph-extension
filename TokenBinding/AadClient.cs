// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace TokenBinding
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    /// <summary>
    /// Helper class for calling onto ADAL
    /// </summary>
    internal class AadClient : IAadClient
    {
        private readonly AuthenticationContext _authContext;
        private readonly ClientCredential _clientCredentials;

        public AadClient(ClientCredential credentials)
        {
            string aadAuthLoginUrl = Constants.DefaultEnvironmentBaseUrl + Constants.DefaultTenantId;
            _authContext = new AuthenticationContext(aadAuthLoginUrl);
            _clientCredentials = credentials;
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
            AuthenticationResult ar = await _authContext.AcquireTokenAsync(resource, _clientCredentials, userAssertion);
            return ar.AccessToken;
        }

        public async Task<string> GetTokenFromClientCredentials(string resource)
        {
            if (string.IsNullOrEmpty(resource))
            {
                throw new ArgumentException("A resource is required to retrieve a token from client credentials.");
            }

            AuthenticationResult authResult = await _authContext.AcquireTokenAsync(resource, _clientCredentials);
            return authResult.AccessToken;
        }
    }
}
