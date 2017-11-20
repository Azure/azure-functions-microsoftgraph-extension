// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    internal class GraphWebhookSubscriptionConverters
    {
        internal class GraphWebhookSubscriptionConverter :
            IAsyncConverter<GraphWebhookSubscriptionAttribute, Subscription[]>,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, string[]>,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, JArray>
        {
            private readonly ServiceManager _serviceManager;
            private readonly GraphWebhookConfig _webhookConfig;

            public GraphWebhookSubscriptionConverter(ServiceManager serviceManager, GraphWebhookConfig webhookConfig)
            {
                _serviceManager = serviceManager;
                _webhookConfig = webhookConfig;
            }

            async Task<Subscription[]> IAsyncConverter<GraphWebhookSubscriptionAttribute, Subscription[]>.ConvertAsync(GraphWebhookSubscriptionAttribute input, CancellationToken cancellationToken)
            {
                return (await GetSubscriptionsFromAttribute(input)).Select(entry => entry.Subscription).ToArray();
            }

            async Task<string[]> IAsyncConverter<GraphWebhookSubscriptionAttribute, string[]>.ConvertAsync(GraphWebhookSubscriptionAttribute input, CancellationToken cancellationToken)
            {
                Subscription[] subscriptions = (await GetSubscriptionsFromAttribute(input)).Select(entry => entry.Subscription).ToArray();
                return subscriptions.Select(sub => sub.Id).ToArray();
            }

            async Task<JArray> IAsyncConverter<GraphWebhookSubscriptionAttribute, JArray>.ConvertAsync(GraphWebhookSubscriptionAttribute input, CancellationToken cancellationToken)
            {
                SubscriptionEntry[] subscriptions = await GetSubscriptionsFromAttribute(input);
                var serializedSubscriptions = new JArray();
                foreach (var subscription in subscriptions)
                {
                    serializedSubscriptions.Add(JObject.FromObject(subscription));
                }
                return serializedSubscriptions;
            }

            protected async Task<SubscriptionEntry[]> GetSubscriptionsFromAttribute(GraphWebhookSubscriptionAttribute attribute)
            {
                IEnumerable<SubscriptionEntry> subscriptionEntries = await _webhookConfig.SubscriptionStore.GetAllSubscriptionsAsync();
                if (TokenIdentityMode.UserFromRequest.ToString().Equals(attribute.Filter, StringComparison.OrdinalIgnoreCase))
                {
                    var dummyTokenAttribute = new TokenAttribute()
                    {
                        Resource = O365Constants.GraphBaseUrl,
                        Identity = TokenIdentityMode.UserFromToken,
                        UserToken = attribute.UserToken,
                        IdentityProvider = "AAD",
                    };
                    var graph = await _serviceManager.GetMSGraphClientAsync(dummyTokenAttribute);
                    var user = await graph.Me.Request().GetAsync();
                    subscriptionEntries = subscriptionEntries.Where(entry => entry.UserId.Equals(user.Id));
                }
                else if (attribute.Filter != null)
                {
                    throw new InvalidOperationException($"There is no filter for {attribute.Filter}");
                }
                return subscriptionEntries.ToArray();
            }
        }

        internal class GenericGraphWebhookSubscriptionConverter<T> : GraphWebhookSubscriptionConverter,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, T[]>
        {
            public GenericGraphWebhookSubscriptionConverter(ServiceManager serviceManager, GraphWebhookConfig webhookConfig) : base(serviceManager, webhookConfig)
            {
            }

            public async Task<T[]> ConvertAsync(GraphWebhookSubscriptionAttribute input, CancellationToken cancellationToken)
            {
                return ConvertSubscriptionEntries(await this.GetSubscriptionsFromAttribute(input));
            }

            //Converts a Subscription Entry into a "flattened" POCO representation where the properties 
            //of the POCO can be UserId or any of the properties of Subscription
            public T[] ConvertSubscriptionEntries(SubscriptionEntry[] entries)
            {
                var pocoType = typeof(T);
                var subEntryType = typeof(Subscription);
                var subscriptionProperties = subEntryType.GetProperties();
                var pocoProperties = pocoType.GetProperties();

                T[] pocos = new T[entries.Length];
                for(int i = 0; i < pocos.Length; i++)
                {
                    pocos[i] = (T)Activator.CreateInstance(typeof(T), new object[] { });
                }

                foreach (PropertyInfo pocoProperty in pocoProperties)
                {
                    var subscriptionProperty = subEntryType.GetProperty(pocoProperty.Name, pocoProperty.PropertyType);
                    if(subscriptionProperty != null)
                    {
                        for(int i = 0; i < pocos.Length; i++)
                        {
                            pocoProperty.SetValue(pocos[i], subscriptionProperty.GetValue(entries[i].Subscription));
                        }
                    }
                }

                var pocoUserIdProperty = pocoType.GetProperty("UserId");
                if (pocoUserIdProperty != null)
                {
                    var subEntryUserIdProperty = typeof(SubscriptionEntry).GetProperty("UserId");
                    for (int i = 0; i < pocos.Length; i++)
                    {
                        pocoUserIdProperty.SetValue(pocos[i], subEntryUserIdProperty.GetValue(entries[i]));
                    }
                }

                return pocos;
            }
        }
    }
}
