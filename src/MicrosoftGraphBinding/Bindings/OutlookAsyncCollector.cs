// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.ObjectModel;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    internal class OutlookAsyncCollector : IAsyncCollector<Message>
    {
        private readonly OutlookService _client;
        private readonly OutlookAttribute _attribute;
        private readonly Collection<Message> _messages;

        public OutlookAsyncCollector(OutlookService client, OutlookAttribute attribute)
        {
            _client = client;
            _attribute = attribute;
            _messages = new Collection<Message>();
        }

        public Task AddAsync(Message item, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (item == null)
            {
                throw new ArgumentNullException("No message item");
            }

            _messages.Add(item);
            return Task.CompletedTask;
        }

        public async Task FlushAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            foreach (var msg in _messages)
            {
                await _client.SendMessageAsync(_attribute, msg, cancellationToken);
            }

            _messages.Clear();
        }
    }
}
