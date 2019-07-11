// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
namespace GraphExtensionSamples
{
    using System;
    using Microsoft.Azure.WebJobs;

    public static class WebhookSubscriptionExamples
    {
        //Note that this function will only pass if the ClientCredentials identity has permission to refresh all of the available webhooks
        public static void RefreshAllSubscriptions(
            [GraphWebhookSubscription] string[] subIds, 
            [GraphWebhookSubscription(
            Action = GraphWebhookSubscriptionAction.Refresh,
            Identity = TokenIdentityMode.ClientCredentials)] ICollector<string> refreshedSubscriptions)
        {
            foreach (var subId in subIds)
            {
                refreshedSubscriptions.Add(subId);
            }
        }

        //This more complicated function without setting up an application identity with lots of permissions
        public static async void RefreshAllSubscriptionsWithoutClientCredentials(
            [GraphWebhookSubscription] SubscriptionPoco[] existingSubscriptions,
            IBinder binder)
        {
            foreach (var subscription in existingSubscriptions)
            {
                // binding in code to allow dynamic identity
                var subscriptionsToRefresh = await binder.BindAsync<IAsyncCollector<string>>(
                    new GraphWebhookSubscriptionAttribute()
                    {
                        Action = GraphWebhookSubscriptionAction.Refresh,
                        Identity = TokenIdentityMode.UserFromRequest
                    }
                );
                {
                    await subscriptionsToRefresh.AddAsync(subscription.Id);
                }
            }
        }

        public static void RefreshSubscriptionsSelectively(
            [GraphWebhookSubscription] SubscriptionPoco[] existingSubscriptions,
            [GraphWebhookSubscription(
            Action = GraphWebhookSubscriptionAction.Refresh,
            Identity = TokenIdentityMode.ClientCredentials)] ICollector<string> refreshedSubscriptions)
        {
            foreach (var subscription in existingSubscriptions)
            {
                if(subscription.ChangeType.Equals("updated") && subscription.ODataType.Equals("#Microsoft.Graph.Message"))
                {
                    refreshedSubscriptions.Add(subscription.Id);
                }   
            }
        }

        public static void DeleteAllSubscriptions(
        [GraphWebhookSubscription(Identity = TokenIdentityMode.ClientCredentials)] string[] subIds,
        [GraphWebhookSubscription(
            Action = GraphWebhookSubscriptionAction.Delete,
            Identity = TokenIdentityMode.ClientCredentials)] ICollector<string> deletedSubscriptions)
        {
            foreach (var subId in subIds)
            {
                deletedSubscriptions.Add(subId);
            }
        }

        public static void SubscribeToInbox([GraphWebhookSubscription(
            Identity = TokenIdentityMode.UserFromRequest,
            SubscriptionResource = "me/mailFolders('Inbox')/messages",
            ChangeTypes = new GraphWebhookChangeType[] {GraphWebhookChangeType.Created, GraphWebhookChangeType.Updated },
            Action = GraphWebhookSubscriptionAction.Create)] out string clientState)
        {
            clientState = Guid.NewGuid().ToString();
        }

        public class SubscriptionPoco
        {
            public string UserId { get; set; }
            //All of the below are properties of the Subscription object that can be used in your custom POCO
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/81c50e72166152f9f84dc38b2516379b7a536300/src/Microsoft.Graph/Models/Generated/Subscription.cs
            //for usage
            public string Id { get; set; }
            public string ODataType { get; set; }
            public string Resource { get; set; }
            public string ChangeType { get; set; }
            public string ClientState { get; set; }
            public string NotificationUrl { get; set; }
            public DateTimeOffset? ExpirationDateTime { get; set; }
        }
        
    }
}
