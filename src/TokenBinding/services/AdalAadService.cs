// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    /// <summary>
    /// Helper class for calling onto ADAL
    /// </summary>
    internal class AdalAadService : IAadService
    {
        private AuthenticationContext _authContext;
        private ClientCredential _clientCredential;

        public AdalAadService(AuthenticationContext authContext, ClientCredential clientCredential)
        {
            _authContext = authContext;
            _clientCredential = clientCredential;
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
            AuthenticationResult ar = await _authContext.AcquireTokenAsync(resource, _clientCredential, userAssertion);
            return ar.AccessToken;
        }

        public async Task<string> GetTokenFromClientCredentials(string resource)
        {
            if (string.IsNullOrEmpty(resource))
            {
                throw new ArgumentException("A resource is required to retrieve a token from client credentials.");
            }

            AuthenticationResult authResult = await _authContext.AcquireTokenAsync(resource, _clientCredential);
            return authResult.AccessToken;
        }
    }
}
