// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
using Microsoft.Azure.WebJobs.Hosting;

[assembly: WebJobsStartup(typeof(AuthTokenWebJobsStartup))]
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    public class AuthTokenWebJobsStartup : IWebJobsStartup
    {
        public void Configure(IWebJobsBuilder builder)
        {
            builder.AddAuthToken();
        }
    }
}
