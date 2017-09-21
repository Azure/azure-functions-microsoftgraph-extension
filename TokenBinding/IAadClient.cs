// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System.Threading.Tasks;

    internal interface IAadClient
    {
        Task<string> GetTokenOnBehalfOfUserAsync(string userToken, string resource);

        Task<string> GetTokenFromClientCredentials(string resource);
    }
}
