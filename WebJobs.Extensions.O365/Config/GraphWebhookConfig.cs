// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;
using System.Threading.Tasks;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    internal class GraphWebhookConfig
    {
        public readonly WebhookSubscriptionStore SubscriptionStore;
        public readonly Uri NotificationUrl;

        private WebhookTriggerBindingProvider _webhookTriggerProvider;

        public GraphWebhookConfig(Uri notificationUrl, WebhookSubscriptionStore subscriptionStore, WebhookTriggerBindingProvider provider)
        {
            SubscriptionStore = subscriptionStore;
            NotificationUrl = notificationUrl;
            _webhookTriggerProvider = provider;
        }

        /// <summary>
        /// Upon receiving webhook trigger data, process it
        /// </summary>
        /// <param name="data">Data from MS Graph -> triggers webhook fx</param>
        /// <returns>Task awaiting result of pushing webhook data</returns>
        internal async Task OnWebhookReceived(WebhookTriggerData data)
        {
            await this._webhookTriggerProvider.PushDataAsync(data);
        }
    }
}
