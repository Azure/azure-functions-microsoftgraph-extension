// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    /// <summary>
    /// List of constants needed by token binding
    /// </summary>
    internal static class Constants
    {
        /// <summary>
        /// Application setting key for the client ID
        /// </summary>
        public const string AppSettingClientIdName = "WEBSITE_AUTH_CLIENT_ID";

        /// <summary>
        /// Application setting key for the client secret
        /// </summary>
        public const string AppSettingClientSecretName = "WEBSITE_AUTH_CLIENT_SECRET";

        /// <summary>
        /// Application setting key for the website hostname
        /// </summary>
        public const string AppSettingWebsiteHostname = "WEBSITE_HOSTNAME";

        /// <summary>
        /// Application setting key for the website auth signing key
        /// </summary>
        public const string AppSettingWebsiteAuthSigningKey = "WEBSITE_AUTH_SIGNING_KEY";

        /// <summary>
        /// The default base url to grab the token from.
        /// </summary>
        public const string DefaultEnvironmentBaseUrl = "https://login.windows.net/";

        /// <summary>
        /// The default tenant to grab the token for
        /// </summary>
        public const string DefaultTenantId = "common";

    }
}
