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
        //The key is the subscription id of the subscription entry
        private IDictionary<string, SubscriptionEntry> _map = new Dictionary<string, SubscriptionEntry>();

        public async Task DeleteAsync(string subscriptionId)
        {
            if (_map.ContainsKey(subscriptionId))
            {
                _map.Remove(subscriptionId);
            }
        }

        public async Task<SubscriptionEntry[]> GetAllSubscriptionsAsync()
        {
            return _map.Values.ToArray();
        }

        public async Task<SubscriptionEntry> GetSubscriptionEntryAsync(string subscriptionId)
        {
            SubscriptionEntry value = null;
            _map.TryGetValue(subscriptionId, out value);
            return value;
        }

        public async Task SaveSubscriptionEntryAsync(Subscription subscription, string userId)
        {
            _map[subscription.Id] = new SubscriptionEntry()
            {
                Subscription = subscription,
                UserId = userId,
            };
        }
    }
}
