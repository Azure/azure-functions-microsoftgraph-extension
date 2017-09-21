// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.Runtime.Serialization;

    /// <summary>
    /// Class representing a single entry in an application's [EasyAuth] Token Store
    /// Names of the fields match the names of the fields returned by the /.auth/me endpoint
    /// </summary>
    [DataContract]
    internal class EasyAuthTokenStoreEntry
    {
        [DataMember(Name = "access_token", EmitDefaultValue = false)]
        public string AccessToken { get; set; }

        [DataMember(Name = "id_token", EmitDefaultValue = false)]
        public string IdToken { get; set; }

        [DataMember(Name = "refresh_token", EmitDefaultValue = false)]
        public string RefreshToken { get; set; }

        [DataMember(Name = "provider_name")]
        public string ProviderName { get; set; }

        [DataMember(Name = "user_id")]
        public string UserId { get; set; }

        [DataMember(Name = "expires_on", EmitDefaultValue = false)]
        public DateTime ExpiresOn { get; set; }

        [DataContract]
        public class Claim
        {
            [DataMember(Name = "typ")]
            public string Type { get; set; }

            [DataMember(Name = "val")]
            public string Value { get; set; }
        }

        [DataMember(Name = "user_claims", EmitDefaultValue = false)]
        public Claim[] UserClaims { get; set; }
    }
}