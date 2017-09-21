// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    /// <summary>
    /// How the binding should grab the access token
    /// </summary>
    public enum IdentityMode
    {
        /// <summary>Same as <see cref="UserFromToken"/> with the user token taken from the X-MS-TOKEN-AAD-ID-TOKEN header. Only works for HttpTrigger</summary>
        UserFromRequest,
        /// <summary> Obtains the access token on behalf of the user whose token is in the userToken field of metadata. </summary>
        UserFromToken,
        /// <summary> Obtains the access token for the user with the id found in the userId field of metadata. </summary>
        UserFromId,
        /// <summary> Obtains the access token for the client credentials found in the application settings. </summary>
        ClientCredentials,
    }
}
