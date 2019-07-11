// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    public class TokenOptions
    {
        private AuthenticationContext _authContext;
        private ClientCredential _clientCredentials;

        /// <summary>
        /// The website hostname
        /// </summary>
        public string HostName { get; set; }

        /// <summary>
        /// Application setting key for the client ID
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// The client secret used with AAD
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Base tenant URL for AAD
        /// </summary>
        public string TenantUrl { get; set; }

        /// <summary>
        /// The signing key used for EasyAuth tokens
        /// </summary>
        public string SigningKey { get; set; }

        /// <summary>
        /// The default base url to gr.
        /// </summary>
        public string DefaultEnvironmentBaseUrl { get; set; } = "https://login.windows.net/";

        /// <summary>
        /// The default tenant to grab the token for
        /// </summary>
        public string DefaultTenantId { get; set; } = "common";
    }
}
