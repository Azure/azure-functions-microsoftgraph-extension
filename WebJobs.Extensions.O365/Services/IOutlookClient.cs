// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System.Threading.Tasks;
    using Microsoft.Graph;
    internal interface IOutlookClient
    {
        Task SendMessageAsync(Message msg);
    }
}
