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
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    /// <summary>
    /// Collector class used to accumulate client states and create new subscriptions and cache them
    /// </summary>
    internal class GraphWebhookSubscriptionAsyncCollector : IAsyncCollector<string>
    {
        private readonly GraphServiceClientManager _clientManager; // already has token
        private readonly ILogger _log;
        private readonly IGraphSubscriptionStore _subscriptionStore;
        private readonly Uri _notificationUrl;
        private readonly GraphOptions _options;
        

        // User attribute that we're bound against.
        // Has key properties (e.g. what resource we're listening to)
        private readonly GraphWebhookSubscriptionAttribute _attribute;

        private List<string> _values;

        public GraphWebhookSubscriptionAsyncCollector(GraphServiceClientManager clientManager, GraphOptions options, ILoggerFactory logFactory, IGraphSubscriptionStore subscriptionStore, Uri notificationUrl, GraphWebhookSubscriptionAttribute attribute)
        {
            _clientManager = clientManager;
            _log = logFactory?.CreateLogger(MicrosoftGraphExtensionConfigProvider.CreateBindingCategory("GraphWebhook"));
            _subscriptionStore = subscriptionStore;
            _notificationUrl = notificationUrl;
            _attribute = attribute;
            _options = options;
            _values = new List<string>();

            _attribute.Validate();
        }

        public Task AddAsync(string value, CancellationToken cancellationToken)
        {
            _values.Add(value);
            return Task.CompletedTask;
        }

        public async Task FlushAsync(CancellationToken cancellationToken)
        {
            switch (_attribute.Action)
            {
                case GraphWebhookSubscriptionAction.Create:
                    await CreateSubscriptionsFromClientStatesAsync(cancellationToken);
                    break;
                case GraphWebhookSubscriptionAction.Delete:
                    await DeleteSubscriptionsFromSubscriptionIds(cancellationToken);
                    break;
                case GraphWebhookSubscriptionAction.Refresh:
                    await RefreshSubscriptionsFromSubscriptionIds(cancellationToken);
                    break;
            }

            _values.Clear();
        }

        private async Task CreateSubscriptionsFromClientStatesAsync(CancellationToken cancellationToken)
        {
            var client = await _clientManager.GetMSGraphClientFromTokenAttributeAsync(_attribute, cancellationToken);
            var userInfo = await client.Me.Request().Select("Id").GetAsync();

            var subscriptions = _values.Select(GetSubscription);
            foreach (var subscription in subscriptions)
            {
                _log.LogTrace($"Sending a request to {_notificationUrl} expecting a 200 response for a subscription to {subscription.Resource}");
                var newSubscription = await client.Subscriptions.Request().AddAsync(subscription);
                await _subscriptionStore.SaveSubscriptionEntryAsync(newSubscription, userInfo.Id);
            }
        }

        private Subscription GetSubscription(string clientState)
        {
            clientState = clientState ?? Guid.NewGuid().ToString();
            return new Subscription
            {
                Resource = _attribute.SubscriptionResource,
                ChangeType = ChangeTypeExtension.ConvertArrayToString(_attribute.ChangeTypes),
                NotificationUrl = _notificationUrl.AbsoluteUri,
                ExpirationDateTime = DateTime.UtcNow + _options.WebhookExpirationTimeSpan,
                ClientState = clientState,
            };
        }

        private async Task DeleteSubscriptionsFromSubscriptionIds(CancellationToken token)
        {
            var client = await _clientManager.GetMSGraphClientFromTokenAttributeAsync(_attribute, token);
            var subscriptionIds = _values;

            foreach (string id in subscriptionIds)
            {
                await DeleteSubscription(client, id);
            }
        }

        private async Task DeleteSubscription(IGraphServiceClient client, string id)
        {
            try
            {
                await client.Subscriptions[id].Request().DeleteAsync();
                _log.LogInformation($"Successfully deleted MS Graph subscription {id}.");
            }
            catch
            {
                _log.LogWarning($"Failed to delete MS Graph subscription {id}.\n Either it never existed or it has already expired.");
            }
            finally
            {
                // Regardless of whether or not deleting the Graph subscription succeeded, delete the file
                await _subscriptionStore.DeleteAsync(id);
            }
        }

        private async Task RefreshSubscriptionsFromSubscriptionIds(CancellationToken token)
        {
            var client = await _clientManager.GetMSGraphClientFromTokenAttributeAsync(_attribute, token);
            var subscriptionIds = _values;

            foreach (var id in subscriptionIds)
            {
                await RefreshSubscription(client, id);
            }
        }

        private async Task RefreshSubscription(IGraphServiceClient client, string id)
        {
            try
            {
                var subscription = new Subscription
                {
                    ExpirationDateTime = DateTime.UtcNow + _options.WebhookExpirationTimeSpan,
                };

                var result = await client.Subscriptions[id].Request().UpdateAsync(subscription);

                _log.LogInformation($"Successfully renewed MS Graph subscription {id}. \n Active until {subscription.ExpirationDateTime}");
            }
            catch (Exception ex)
            {
                // If the subscription is expired, it can no longer be renewed, so delete the file
                var subscriptionEntry = await _subscriptionStore.GetSubscriptionEntryAsync(id);
                if (subscriptionEntry != null)
                {
                    if(subscriptionEntry.Subscription.ExpirationDateTime < DateTime.UtcNow)
                    {
                        await _subscriptionStore.DeleteAsync(id);
                    } else
                    {
                        _log.LogError(ex, "A non-expired subscription failed to renew");
                    }
                } else
                {
                    _log.LogWarning("The subscription with id " + id + " was not present in the local subscription store.");
                }
            }
        }
    }
}
