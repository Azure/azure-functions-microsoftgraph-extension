// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace TokenBinding
{
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    public class AadClientFactory
    {
        public virtual IAadClient GetClient(ClientCredential credentials)
        {
            return new AadClient(credentials);
        }
    }
}
