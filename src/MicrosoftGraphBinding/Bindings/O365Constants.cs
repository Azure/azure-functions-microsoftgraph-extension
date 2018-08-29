// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;

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
    }
}
