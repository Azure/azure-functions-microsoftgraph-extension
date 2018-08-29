// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph;
using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
using Microsoft.Azure.WebJobs.Hosting;

[assembly: WebJobsStartup(typeof(MicrosoftGraphWebJobsStartup))]
namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    public class MicrosoftGraphWebJobsStartup : IWebJobsStartup
    {
        public void Configure(IWebJobsBuilder builder)
        {
            builder.AddMicrosoftGraph();
        }
    }
}
