// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    internal static class Constants
    {
        public const string ClientIdName = "WEBSITE_AUTH_CLIENT_ID";

        public const string ClientSecretName = "WEBSITE_AUTH_CLIENT_SECRET";

        public const string WebsiteHostname = "WEBSITE_HOSTNAME";

        public const string WebsiteAuthSigningKey = "WEBSITE_AUTH_SIGNING_KEY";

        public const string WebsiteAuthOpenIdIssuer = "WEBSITE_AUTH_OPENID_ISSUER";

        public const string EasyAuthAadAccessTokenHeader = "X-MS-TOKEN-AAD-ID-TOKEN";
    }
}
