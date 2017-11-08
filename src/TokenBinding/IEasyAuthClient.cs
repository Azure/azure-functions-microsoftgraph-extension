// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System.Threading.Tasks;

    internal interface IEasyAuthClient
    {
        Task<EasyAuthTokenStoreEntry> GetTokenStoreEntry(TokenBaseAttribute attribute);

        Task RefreshToken(TokenBaseAttribute attribute);
    }
}
