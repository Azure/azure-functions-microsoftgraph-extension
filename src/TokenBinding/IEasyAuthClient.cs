// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System.IdentityModel.Tokens.Jwt;
    using System.Threading.Tasks;

    public interface IEasyAuthClient
    {
        Task<EasyAuthTokenStoreEntry> GetTokenStoreEntry(JwtSecurityToken jwt, TokenBaseAttribute attribute);

        Task RefreshToken(JwtSecurityToken jwt, TokenBaseAttribute attribute);
    }
}
