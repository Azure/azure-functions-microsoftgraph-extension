// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    internal class OutlookService
    {
        private IGraphServiceClient _client;

        public OutlookService(IGraphServiceClient client)
        {
            _client = client;
        }

        public async Task SendMessageAsync(Message msg)
        {
            await _client.SendMessageAsync(msg);
        }
    }
}
