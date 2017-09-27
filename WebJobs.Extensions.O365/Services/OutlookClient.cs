// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    internal class OutlookClient : IOutlookClient
    {
        private Task<IGraphServiceClient> _client;

        public OutlookClient(Task<IGraphServiceClient> client)
        {
            _client = client;
        }

        /// <summary>
        /// Send an email with a dynamically set body
        /// </summary>
        /// <param name="client">GraphServiceClient used to send request</param>
        /// <returns>Async task for posted message</returns>
        public async Task SendMessageAsync(Message msg)
        {
            // Send message & save to sent items folder
            await (await _client).Me.SendMail(msg, true).Request().PostAsync();
        }
    }
}
