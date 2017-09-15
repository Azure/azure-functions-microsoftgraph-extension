// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using System;
    using System.Collections.ObjectModel;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;

    internal class OutlookAsyncCollector : IAsyncCollector<Message>
    {
        private readonly GraphServiceClient client;
        private readonly OutlookAttribute attribute;
        private readonly Collection<Message> messages = new Collection<Message>();

        public OutlookAsyncCollector(GraphServiceClient client, OutlookAttribute attribute)
        {
            this.client = client;
            this.attribute = attribute;
        }

        public Task AddAsync(Message item, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (item == null)
            {
                throw new ArgumentNullException("No message item");
            }

            this.messages.Add(item);
            return Task.CompletedTask;
        }

        public async Task FlushAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            foreach (var msg in this.messages)
            {
                await this.client.SendMessage(this.attribute, msg);
            }

            this.messages.Clear();
        }
    }
}
