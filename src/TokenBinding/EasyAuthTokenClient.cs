// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Net;
    using System.Net.Http;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Newtonsoft.Json;

    /// <summary>
    /// The client responsible for handling all EasyAuth token-related tasks.
    /// </summary>
    internal class EasyAuthTokenClient : IEasyAuthClient
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private readonly ILogger _log;
        private readonly TokenOptions _options;

        /// <summary>
        /// Initializes a new instance of the <see cref="EasyAuthTokenClient"/> class.
        /// </summary>
        /// <param name="hostName">The hostname of the webapp </param>
        /// <param name="signingKey">The website authorization signing key</param>
        public EasyAuthTokenClient(IOptions<TokenOptions> options, ILoggerFactory loggerFactory)
        {
            _log = loggerFactory.CreateLogger(AuthTokenExtensionConfigProvider.CreateBindingCategory("AuthToken"));
            _options = options.Value;
        }

        public async Task<EasyAuthTokenStoreEntry> GetTokenStoreEntry(JwtSecurityToken jwt, TokenBaseAttribute attribute)
        {
            string httpScheme = "https";

            if (_options.HostName.StartsWith("localhost:") || _options.HostName.StartsWith("127.0.0.1:"))
            {
                httpScheme = "http";
            }

            // Send the token to the local /.auth/me endpoint and return the JSON
            string meUrl = $"{httpScheme}://{_options.HostName}/.auth/me?provider={attribute.IdentityProvider}";

            using (var request = new HttpRequestMessage(HttpMethod.Get, meUrl))
            {
                request.Headers.Add("x-zumo-auth", jwt.RawData);
                _log.LogTrace($"Fetching user token data from {meUrl}");
                using (HttpResponseMessage response = await _httpClient.SendAsync(request))
                {
                    _log.LogTrace($"Response from '${meUrl}: {response.StatusCode}");
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorResponse = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException($"Request to {meUrl} failed. Status Code: {response.StatusCode}; Body: {errorResponse}");
                    }
                    var responseString = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<EasyAuthTokenStoreEntry>(responseString);
                }
            }
        }

        public async Task RefreshToken(JwtSecurityToken jwt, TokenBaseAttribute attribute)
        {
            string httpScheme = "https";

            if (_options.HostName.StartsWith("localhost:") || _options.HostName.StartsWith("127.0.0.1:"))
            {
                httpScheme = "http";
            }

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

            string refreshUrl = $"{httpScheme}://{_options.HostName}/.auth/refresh?resource=" + WebUtility.UrlEncode(attribute.Resource);

            using (var refreshRequest = new HttpRequestMessage(HttpMethod.Get, refreshUrl))
            {
                refreshRequest.Headers.Add("x-zumo-auth", jwt.RawData);
                _log.LogTrace($"Refreshing ${attribute.IdentityProvider} access token for user ${attribute.UserId} at ${refreshUrl}");

                using (HttpResponseMessage response = await _httpClient.SendAsync(refreshRequest))
                {
                    _log.LogTrace($"Response from ${refreshUrl}: {response.StatusCode}");
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
