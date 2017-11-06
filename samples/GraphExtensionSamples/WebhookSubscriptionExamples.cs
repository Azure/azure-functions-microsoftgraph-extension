// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
namespace GraphExtensionSamples
{
    using System;
    using Microsoft.Azure.WebJobs;

    public static class WebhookSubscriptionExamples
    {
        public static void RefreshAllSubscriptions(
            [GraphWebhookSubscription(Identity = TokenIdentityMode.ClientCredentials)] string[] subIds,
            [GraphWebhookSubscription(
            Action = GraphWebhookSubscriptionAction.Refresh,
            Identity = TokenIdentityMode.ClientCredentials)] ICollector<string> refreshedSubscriptions)
        {
            foreach (var subId in subIds)
            {
                refreshedSubscriptions.Add(subId);
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
            UserId = "%UserId%",
            Identity = TokenIdentityMode.UserFromId,
            SubscriptionResource = "me/mailFolders('Inbox')/messages",
            ChangeTypes = new GraphWebhookChangeType[] {GraphWebhookChangeType.Created, GraphWebhookChangeType.Updated },
            Action = GraphWebhookSubscriptionAction.Create)] out string clientState)
        {
            clientState = Guid.NewGuid().ToString();
        }
    }
}
