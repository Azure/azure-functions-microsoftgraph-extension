// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Net;
    using System.Net.Http;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Host;
    using Newtonsoft.Json;

    /// <summary>
    /// The client responsible for handling all EasyAuth token-related tasks.
    /// </summary>
    internal class EasyAuthTokenClient : IEasyAuthClient
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        private readonly string _baseUrl;

        private readonly TraceWriter _log;

        private JwtSecurityToken _tokenForEasyAuthAccess;

        /// <summary>
        /// Initializes a new instance of the <see cref="EasyAuthTokenClient"/> class.
        /// </summary>
        /// <param name="hostName">The hostname of the webapp </param>
        /// <param name="signingKey">The website authorization signing key</param>
        public EasyAuthTokenClient(string hostName, TraceWriter log)
        {
            _baseUrl = "https://" + hostName + "/";
            _log = log;
        }

        public string GetBaseUrl()
        {
            return _baseUrl;
        }

        public async Task<EasyAuthTokenStoreEntry> GetTokenStoreEntry(JwtSecurityToken jwt, TokenAttribute attribute)
        {
            // Send the token to the local /.auth/me endpoint and return the JSON
            string meUrl = _baseUrl + $".auth/me?provider={attribute.IdentityProvider}";

            using (var request = new HttpRequestMessage(HttpMethod.Get, meUrl))
            {
                request.Headers.Add("x-zumo-auth", jwt.RawData);
                _log.Verbose($"Fetching user token data from {meUrl}");
                using (HttpResponseMessage response = await _httpClient.SendAsync(request))
                {
                    _log.Verbose($"Response from '${meUrl}: {response.StatusCode}");
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorResponse = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException($"Request to {_baseUrl} failed. Status Code: {response.StatusCode}; Body: {errorResponse}");
                    }
                    var responseString = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<EasyAuthTokenStoreEntry>(responseString);
                }
            }
        }

        public async Task RefreshToken(JwtSecurityToken jwt, TokenAttribute attribute)
        {
            if (string.IsNullOrEmpty(attribute.Resource))
            {
                throw new ArgumentException("A resource is required to renew an access token.");
            }

            if (string.IsNullOrEmpty(attribute.UserId))
            {
                throw new ArgumentException("A userId is required to renew an access token.");
            }

            if (string.IsNullOrEmpty(attribute.IdentityProvider))
            {
                throw new ArgumentException("A provider is necessary to renew an access token.");
            }

            string refreshUrl = _baseUrl + $".auth/refresh?resource=" + WebUtility.UrlEncode(attribute.Resource);

            using (var refreshRequest = new HttpRequestMessage(HttpMethod.Get, refreshUrl))
            {
                refreshRequest.Headers.Add("x-zumo-auth", jwt.RawData);
                _log.Verbose($"Refreshing ${attribute.IdentityProvider} access token for user ${attribute.UserId} at ${refreshUrl}");
                using (HttpResponseMessage response = await _httpClient.SendAsync(refreshRequest))
                {
                    _log.Verbose($"Response from ${refreshUrl}: {response.StatusCode}");
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorResponse = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException($"Failed to refresh {attribute.UserId} {attribute.IdentityProvider} error={response.StatusCode} {errorResponse}");
                    }
                }
            }
        }
    }
}
