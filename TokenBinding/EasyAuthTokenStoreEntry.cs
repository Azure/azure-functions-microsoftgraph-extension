// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.Token.Tests")]
namespace TokenBinding
{
    using System;

    /// <summary>
    /// Class representing a single entry in an application's [EasyAuth] Token Store
    /// Names of the fields match the names of the fields returned by the /.auth/me endpoint
    /// </summary>
    internal class EasyAuthTokenStoreEntry
    {
        public string access_token { get; set; }

        public string id_token { get; set; } // same value as I'd get from X-MS-TOKEN-AAD-ID-TOKEN header

        public string refresh_token { get; set; }

        public string provider_name { get; set; }

        public string user_id { get; set; }

        public DateTime expires_on { get; set; }

        public class Claim
        {
            public string typ { get; set; }

            public string val { get; set; }
        }

        public Claim[] user_claims { get; set; }
    }
}