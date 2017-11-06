// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.Threading.Tasks;

    /// <summary>
    /// Class representing an application's [EasyAuth] Token Store
    /// see  https://cgillum.tech/2016/03/07/app-service-token-store/
    /// </summary>
    internal class EasyAuthTokenManager
    {
        private static readonly int GraphTokenBufferInMinutes = 5;

        private readonly IEasyAuthClient _client;

        /// <summary>
        /// Initializes a new instance of the <see cref="EasyAuthTokenManager"/> class.
        /// </summary>
        /// <param name="hostName">The hostname of the keystore. </param>
        /// <param name="signingKey">The website authorization signing key</param>
        public EasyAuthTokenManager(IEasyAuthClient client)
        {
            _client = client;
        }

        /// <summary>
        /// Retrieve Easy Auth token based on provider & principal ID
        /// </summary>
        /// <param name="attribute">The metadata for the token to grab</param>
        /// <returns>Task with Token Store entry of the token</returns>
        public async Task<string> GetEasyAuthAccessTokenAsync(TokenAttribute attribute)
        {
            EasyAuthTokenStoreEntry tokenStoreEntry = await _client.GetTokenStoreEntry(attribute);

            bool isTokenValid = IsTokenValid(tokenStoreEntry.AccessToken);
            bool isTokenExpired = tokenStoreEntry.ExpiresOn <= DateTime.UtcNow.AddMinutes(GraphTokenBufferInMinutes);
            bool isRefreshable = IsRefreshableProvider(attribute.IdentityProvider);

            if (isRefreshable && (isTokenExpired || !isTokenValid))
            {
                await _client.RefreshToken(attribute);

                // Now that the refresh has occured, grab the new token
                tokenStoreEntry = await _client.GetTokenStoreEntry(attribute);
            }

            return tokenStoreEntry.AccessToken;
        }

        private static bool IsTokenValid(string token)
        {
            return EasyAuthTokenClient.JwtHandler.CanReadToken(token);
        }

        private static bool IsRefreshableProvider(string provider)
        {
            //TODO: For now, since we are focusing on AAD, only include it in the refresh path.
            return provider.Equals("AAD", StringComparison.OrdinalIgnoreCase);
        }
    }
}
