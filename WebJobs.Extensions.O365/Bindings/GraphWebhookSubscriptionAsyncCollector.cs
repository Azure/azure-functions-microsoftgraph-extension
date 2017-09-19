// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using System.Linq;

    /// <summary>
    /// Collector class used to accumulate client states and create new subscriptions and cache them
    /// </summary>
    public class GraphWebhookSubscriptionAsyncCollector : IAsyncCollector<string>
    {
        private readonly O365Extension _extension; // already has token

        // User attribute that we're bound against.
        // Has key properties (e.g. what resource we're listening to)
        private readonly GraphWebhookSubscriptionAttribute _attribute;

        private List<string> _values;

        public GraphWebhookSubscriptionAsyncCollector(O365Extension extension, GraphWebhookSubscriptionAttribute attribute)
        {
            _extension = extension;
            _attribute = attribute;
            _values = new List<string>();

            _attribute.Validate();
        }

        public async Task AddAsync(string value, CancellationToken cancellationToken = default(CancellationToken))
        {
            _values.Add(value);
        }

        public async Task FlushAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            switch (_attribute.Action)
            {
                case GraphWebhookSubscriptionAction.Create:
                    await CreateSubscriptionsFromClientStates();
                    break;
                case GraphWebhookSubscriptionAction.Delete:
                    await DeleteSubscriptionsFromSubscriptionIds();
                    break;
                case GraphWebhookSubscriptionAction.Refresh:
                    await RefreshSubscriptionsFromSubscriptionIds();
                    break;
            }

            _values.Clear();
        }

        private async Task CreateSubscriptionsFromClientStates()
        {
            var client = await _extension.GetMSGraphClientAsync(_attribute);
            var userInfo = await client.Me.Request().Select("Id").GetAsync();
            var cache = _extension.subscriptionStore;

            var subscriptions = _values.Select(GetSubscription);
            foreach (var subscription in subscriptions)
            {
                _extension.Log.Verbose($"Sending a request to {_extension.NotificationUrl} expecting a 200 response for a subscription to {subscription.Resource}");
                var newSubscription = await client.Subscriptions.Request().AddAsync(subscription);
                await cache.SaveSubscriptionEntryAsync(newSubscription, userInfo.Id);
            }
        }

        private Subscription GetSubscription(string clientState)
        {
            clientState = clientState ?? Guid.NewGuid().ToString();
            return new Subscription
            {
                Resource = _attribute.SubscriptionResource,
                ChangeType = ChangeTypeExtension.ConvertArrayToString(_attribute.ChangeTypes),
                NotificationUrl = _extension.NotificationUrl.ToString(),
                ExpirationDateTime = DateTime.UtcNow + O365Constants.WebhookExpirationTimeSpan,
                ClientState = clientState,
            };
        }

        private async Task DeleteSubscriptionsFromSubscriptionIds()
        {
            var client = await _extension.GetMSGraphClientAsync(_attribute);
            var subscriptionIds = _values;

            foreach (string id in subscriptionIds)
            {
                Task.Run(() => DeleteSubscription(client, id));
            }
        }

        private async void DeleteSubscription(GraphServiceClient client, string id)
        {
            try
            {
                await client.Subscriptions[id].Request().DeleteAsync();
                _extension.Log.Info($"Successfully deleted MS Graph subscription {id}.");
            }
            catch
            {
                _extension.Log.Info($"Failed to delete MS Graph subscription {id}.\n Either it never existed or it has already expired.");
            }
            finally
            {
                // Regardless of whether or not deleting the Graph subscription succeeded, delete the file
                _extension.subscriptionStore.DeleteAsync(id);
            }
        }

        private async Task RefreshSubscriptionsFromSubscriptionIds()
        {
            var client = await _extension.GetMSGraphClientAsync(_attribute);
            var subscriptionIds = _values;

            foreach (var id in subscriptionIds)
            {
                Task.Run(() => RefreshSubscription(client, id));
            }
        }

        private async void RefreshSubscription(GraphServiceClient client, string id)
        {
            try
            {
                var subscription = new Subscription
                {
                    ExpirationDateTime = DateTime.UtcNow + O365Constants.WebhookExpirationTimeSpan,
                };

                var result = await client.Subscriptions[id].Request().UpdateAsync(subscription);

                _extension.Log.Info($"Successfully renewed MS Graph subscription {id}. \n Active until {subscription.ExpirationDateTime}");
            }
            catch
            {
                _extension.Log.Info($"Failed to renew MS Graph subscription {id}.\n Either it never existed or it has already expired.");

                // If the subscription is expired, it can no longer be renewed, so delete the file
                _extension.subscriptionStore.DeleteAsync(id);
            }
        }
    }
}
