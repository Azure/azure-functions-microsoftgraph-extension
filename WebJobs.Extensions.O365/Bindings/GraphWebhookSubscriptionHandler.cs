// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Azure.WebJobs.Extensions.Bindings;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    // Handles the subscription validation and notification payloads
    internal class GraphWebhookSubscriptionHandler
    {
        private readonly MicrosoftGraphExtensionConfig _extension; // already has token

        public GraphWebhookSubscriptionHandler(MicrosoftGraphExtensionConfig extension)
        {
            _extension = extension;
        }

        private async Task HandleNotifications(NotificationPayload notifications)
        {
            // A single webhook might get notifications from different users. 
            List<WebhookTriggerData> resources = new List<WebhookTriggerData>();

            var subscriptionStore = _extension.subscriptionStore;

            foreach (Notification notification in notifications.Value)
            {
                var subId = notification.SubscriptionId;
                var entry = await subscriptionStore.GetSubscriptionEntryAsync(subId);
                if (entry == null)
                {
                    _extension.Log.Error($"No subscription exists in our store for subscription id: {subId}");
                    // mapping of subscription ID to principal ID does not exist in file system
                    continue;
                }

                if (entry.Subscription.ClientState != notification.ClientState)
                {
                    _extension.Log.Verbose($"The subscription store's client state: {entry.Subscription.ClientState} did not match the notifications's client state: {notification.ClientState}");
                    // Stored client state does not match client state we just received
                    continue;
                }

                // call onto Graph to fetch the resource
                var userId = entry.UserId;
                var graphClient = await _extension.GetMSGraphClientFromUserIdAsync(userId);

                _extension.Log.Verbose($"A graph client was obtained for subscription id: {subId}");

                // Prepend with / if necessary
                if (notification.Resource[0] != '/')
                {
                    notification.Resource = '/' + notification.Resource;
                }

                var url = graphClient.BaseUrl + notification.Resource;

                HttpRequestMessage request = new HttpRequestMessage
                {
                    Method = HttpMethod.Get,
                    RequestUri = new Uri(url),
                };

                _extension.Log.Verbose($"Making a GET request to {url} on behalf of subId: {subId}");

                await graphClient.AuthenticationProvider.AuthenticateRequestAsync(request); // Add authentication header
                var response = await graphClient.HttpProvider.SendAsync(request);
                string responseContent = await response.Content.ReadAsStringAsync();

                _extension.Log.Verbose($"Recieved {responseContent} from request to {url}");

                var actualPayload = JObject.Parse(responseContent);

                // Superimpose some common properties onto the JObject for easy access.
                actualPayload["ClientState"] = entry.Subscription.ClientState;

                // Drive items only payload without resource data
                string odataType = null;
                if (notification.ResourceData != null)
                {
                    odataType = notification.ResourceData.ODataType;
                }
                else if (notification.Resource.StartsWith("/me/drive"))
                {
                    odataType = "#Microsoft.Graph.DriveItem";
                }

                actualPayload["@odata.type"] = odataType;

                var data = new WebhookTriggerData
                {
                    client = graphClient,
                    Payload = actualPayload,
                    odataType = odataType,
                };

                resources.Add(data);
            }

            _extension.Log.Verbose($"Triggering {resources.Count} GraphWebhookTriggers");
            Task[] webhookReceipts = resources.Select(item => _extension.OnWebhookReceived(item)).ToArray();

            Task.WaitAll(webhookReceipts);
            _extension.Log.Verbose($"Finished responding to notifications.");
        }

        // See here for subscribing and payload information.
        // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/subscription_post_subscriptions
        public async Task<HttpResponseMessage> ProcessAsync(
            HttpRequestMessage request)
        {
            var nvc = HttpUtility.ParseQueryString(request.RequestUri.Query);

            string validationToken = nvc["validationToken"];
            if (validationToken != null)
            {
                
                return HandleInitialValidation(validationToken);
            }

            return await HandleNotificationPayload(request);
        }

        private HttpResponseMessage HandleInitialValidation(string validationToken)
        {
            _extension.Log.Verbose($"Returning a 200 OK Response to a request to {_extension.NotificationUrl} with a validation token of {validationToken}");
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(validationToken, Encoding.UTF8, "plain/text"),
            };
        }

        private async Task<HttpResponseMessage> HandleNotificationPayload(HttpRequestMessage request)
        {
            string json = await request.Content.ReadAsStringAsync();
            var notifications = JsonConvert.DeserializeObject<NotificationPayload>(json);

            _extension.Log.Verbose($"Received a notification payload of {json}");
            // We have 30sec to reply to the payload.
            // So offload everything else (especially fetches back to the graph and executing the user function)
            Task.Run(() => HandleNotifications(notifications));

            // Still return a 200 so the service doesn't resend the notification.
            return new HttpResponseMessage(HttpStatusCode.OK);
        }


    }
}
