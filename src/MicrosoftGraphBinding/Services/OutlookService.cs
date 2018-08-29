// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    internal class OutlookService
    {
        private GraphServiceClientManager _clientManager;

        public OutlookService(GraphServiceClientManager clientManager)
        {
            _clientManager = clientManager;
        }

        public async Task SendMessageAsync(OutlookAttribute attr, Message msg, CancellationToken token)
        {
            IGraphServiceClient client = await _clientManager.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            await client.SendMessageAsync(msg, token);
        }
    }
}
