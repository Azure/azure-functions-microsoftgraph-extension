// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Graph;

    class MemorySubscriptionStore : IGraphSubscriptionStore
    {
        private IDictionary<string, SubscriptionEntry> map;

        public MemorySubscriptionStore()
        {
            map = new Dictionary<string, SubscriptionEntry>();
        }

        public async Task DeleteAsync(string subscriptionId)
        {
            map.Remove(subscriptionId);
        }

        public async Task<SubscriptionEntry[]> GetAllSubscriptionsAsync()
        {
            return map.Select((keyValuePair, index) => keyValuePair.Value).ToArray();
        }

        public async Task<SubscriptionEntry> GetSubscriptionEntryAsync(string subId)
        {
            SubscriptionEntry value = null;
            map.TryGetValue(subId, out value);
            return value;
        }

        public async Task SaveSubscriptionEntryAsync(Subscription subscription, string userId)
        {
            map[subscription.Id] = new SubscriptionEntry()
            {
                Subscription = subscription,
                UserId = userId,
            };
        }
    }
}
