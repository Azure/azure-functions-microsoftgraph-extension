// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace TokenBinding
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Description;
    using System;

    // Requires that a AAD clientId/clientSecret are set. These are from EasyAuth. 
    // Resource - gets the audience/scopes. (appId
    // Modes:
    //    1. From Request - get idToken from request (X-MS-TOKEN-AAD-ID-TOKEN in EA) 
    //        Set: User="Auth" 
    //    2. From EA TokenStore (previously logged in) 
    //        Uses refresh flow
    //        Set UserId = id. 

    /// <summary>
    /// Bind to an AAD token.
    /// Also serves as a base-class for other AAD-bindings.
    /// </summary>
    [Binding]
    public class TokenAttribute : Attribute
    {
        private IdentityMode _identity;

        /// <summary>
        /// Gets or sets a resource for a token exchange. Optional
        /// </summary>
        public string Resource { get; set; }

        /// <summary>
        /// Gets or sets an identity provider for the token exchange. Optional
        /// </summary>
        public string IdentityProvider { get; set; }

        /// <summary>
        /// Gets or sets token to grab on-behalf-of. Required if Identity="userFromToken".
        /// </summary>
        [AutoResolve]
        public string UserToken { get; set; }

        /// <summary>
        /// Gets or sets user id to grab token for. Required if Identity="userFromId".
        /// </summary>
        [AutoResolve]
        public string UserId { get; set; }

        /// <summary>
        /// Gets or sets how to determine identity. Required.
        /// </summary>
        public IdentityMode Identity
        {
            get
            {
                return _identity;
            }

            set
            {
                if (value == IdentityMode.UserFromRequest)
                {
                    _identity = IdentityMode.UserFromToken;
                    this.UserToken = "{headers.X-MS-TOKEN-AAD-ID-TOKEN}";
                }
                else
                {
                    _identity = value;
                }
            }
        }

        public void CheckValidity()
        {
            switch (this.Identity)
            {
                case IdentityMode.ClientCredentials:
                    break;
                case IdentityMode.UserFromId:
                    if (string.IsNullOrWhiteSpace(this.UserId))
                    {
                        throw new FormatException("A token attribute with identity=userFromId requires a userId");
                    }

                    break;
                case IdentityMode.UserFromToken:
                    if (string.IsNullOrWhiteSpace(this.UserToken))
                    {
                        throw new FormatException("A token attribute with identity=userFromToken requires a userToken");
                    }

                    break;
            }
        }
    }
}
