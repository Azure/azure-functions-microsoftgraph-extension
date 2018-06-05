// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    /// <summary>
    /// List of constants needed by token binding
    /// </summary>
    internal static class Constants
    {
        #region Always Present AppSettings
        /// <summary>
        /// Application setting key for the website hostname
        /// </summary>
        public const string AppSettingWebsiteHostname = "WEBSITE_HOSTNAME";
        #endregion

        #region AAD Client AppSettings
        /// <summary>
        /// Application setting key for the client ID
        /// </summary>
        public const string AppSettingClientIdName = "WEBSITE_AUTH_CLIENT_ID";

        /// <summary>
        /// Application setting key for the client secret
        /// </summary>
        public const string AppSettingClientSecretName = "WEBSITE_AUTH_CLIENT_SECRET";

        /// <summary>
        /// Base tenant URL for AAD
        /// </summary>
        public const string AppSettingWebsiteAuthOpenIdIssuer = "WEBSITE_AUTH_OPENID_ISSUER";
        #endregion

        #region EasyAuth Required AppSettings
        /// <summary>
        /// Application setting key for the website auth signing key
        /// </summary>
        public const string AppSettingWebsiteAuthSigningKey = "WEBSITE_AUTH_SIGNING_KEY";
        #endregion

        #region Default Values
        /// <summary>
        /// The default AAD tenant url token.
        /// </summary>
        public const string DefaultAadTenantUrl = "https://login.windows.net/common";
        #endregion
    }
}
