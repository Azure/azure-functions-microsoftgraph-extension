// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Host;
    using Microsoft.Graph;

    /// <summary>
    /// Collector class used to accumulate client states and create new subscriptions and cache them
    /// </summary>
    internal class GraphWebhookSubscriptionAsyncCollector : IAsyncCollector<string>
    {
        private readonly ServiceManager _extension; // already has token
        private readonly TraceWriter _log;
        private readonly GraphWebhookConfig _webhookConfig;

        // User attribute that we're bound against.
        // Has key properties (e.g. what resource we're listening to)
        private readonly GraphWebhookSubscriptionAttribute _attribute;

        private List<string> _values;

        public GraphWebhookSubscriptionAsyncCollector(ServiceManager extension, TraceWriter log, GraphWebhookConfig config, GraphWebhookSubscriptionAttribute attribute)
        {
            _extension = extension;
            _log = log;
            _webhookConfig = config;
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
            var cache = _webhookConfig.SubscriptionStore;

            var subscriptions = _values.Select(GetSubscription);
            foreach (var subscription in subscriptions)
            {
                _log.Verbose($"Sending a request to {_webhookConfig.NotificationUrl} expecting a 200 response for a subscription to {subscription.Resource}");
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
                NotificationUrl = _webhookConfig.NotificationUrl.ToString(),
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

        private async void DeleteSubscription(IGraphServiceClient client, string id)
        {
            try
            {
                await client.Subscriptions[id].Request().DeleteAsync();
                _log.Info($"Successfully deleted MS Graph subscription {id}.");
            }
            catch
            {
                _log.Info($"Failed to delete MS Graph subscription {id}.\n Either it never existed or it has already expired.");
            }
            finally
            {
                // Regardless of whether or not deleting the Graph subscription succeeded, delete the file
                await _webhookConfig.SubscriptionStore.DeleteAsync(id);
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

        private async void RefreshSubscription(IGraphServiceClient client, string id)
        {
            try
            {
                var subscription = new Subscription
                {
                    ExpirationDateTime = DateTime.UtcNow + O365Constants.WebhookExpirationTimeSpan,
                };

                var result = await client.Subscriptions[id].Request().UpdateAsync(subscription);

                _log.Info($"Successfully renewed MS Graph subscription {id}. \n Active until {subscription.ExpirationDateTime}");
            }
            catch (Exception ex)
            {
                // If the subscription is expired, it can no longer be renewed, so delete the file
                var subscriptionEntry = await _webhookConfig.SubscriptionStore.GetSubscriptionEntryAsync(id);
                if (subscriptionEntry != null)
                {
                    if(subscriptionEntry.Subscription.ExpirationDateTime < DateTime.UtcNow)
                    {
                        _webhookConfig.SubscriptionStore.DeleteAsync(id);
                    } else
                    {
                        _log.Error("A non-expired subscription failed to renew", ex);
                    }
                } else
                {
                    _log.Warning("The subscription with id " + id + " was not present in the local subscription store.");
                }
            }
        }
    }
}
