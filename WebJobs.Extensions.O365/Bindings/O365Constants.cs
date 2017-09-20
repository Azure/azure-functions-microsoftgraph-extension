// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;

namespace Microsoft.Azure.WebJobs.Extensions
{
    /// <summary>
    /// Constants needed to communicate with Graph, find app settings, etc.
    /// </summary>
    internal static class O365Constants
    {
        /// <summary>
        /// Base URL used to make any rest API calls
        /// </summary>
        public const string GraphBaseUrl = "https://graph.microsoft.com";

        /// <summary>
        /// App setting key to find application Client ID
        /// </summary>
        public const string AppSettingClientIdName = "WEBSITE_AUTH_CLIENT_ID";

        /// <summary>
        /// App setting key to find application Client Secret
        /// </summary>
        public const string AppSettingClientSecretName = "WEBSITE_AUTH_CLIENT_SECRET";

        /// <summary>
        /// App setting key to find where tokens and subscriptions are stored
        /// </summary>
        public const string AppSettingBYOBTokenMap = "BYOB_TokenMap";

        /// <summary>
        /// If AppSettingBYOBTokenMap's key is not set, use this value
        /// </summary>
        public const string DefaultBYOBLocation = "D:/home/data/byob_graphmap";

        /// <summary>
        /// JObject key that holds range values (I/O)
        /// </summary>
        public const string ValuesKey = "Microsoft.O365Bindings.values";

        /// <summary>
        /// JObject key that holds number of rows in ValuesKey
        /// </summary>
        public const string RowsKey = "Microsoft.O365Bindings.rows";

        /// <summary>
        /// JObject key that holds number of columns in ValuesKey (rectangular data)
        /// </summary>
        public const string ColsKey = "Microsoft.O365Bindings.columns";

        /// <summary>
        /// JObject key to indicate whether or not data came from POCO objects (list, array, or otherwise)
        /// </summary>
        public const string POCOKey = "Microsoft.O365Bindings.POCO";

        /// <summary>
        /// The maximum expiration timespan that does not violate any webhook's maximum expiration timespan.
        /// </summary>
        public static readonly TimeSpan WebhookExpirationTimeSpan = new TimeSpan(0, 0, 4230, 0);
    }
}
