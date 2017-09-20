// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    /// <summary>
    /// GraphServiceClient + time the token used to authenticate calls expires
    /// </summary>
    internal class CachedClient {
        internal GraphServiceClient client;
        internal int expirationDate;
    }

    /// <summary>
    /// Enum of changes that can be subscribed to
    /// </summary>
    public enum ChangeType
    {
        /// <summary>
        /// Webhook activated when a new item of the subscribed resource is CREATED
        /// </summary>
        Created,

        /// <summary>
        /// Webhook activated when a new item of the subscribed resource is UPDATED
        /// </summary>
        Updated,

        /// <summary>
        /// Webhook activated when a new item of the subscribed resource is DELETED
        /// </summary>
        Deleted,
    }

    /// <summary>
    /// Class containing an array of notifications received in a single blast from MS Graph
    /// </summary>
    internal class NotificationPayload
    {
        /// <summary>
        /// Gets or sets the array of notifications received by the function app
        /// </summary>
        public Notification[] Value { get; set; }
    }

    /// <summary>
    /// Single notification from MS Graph indicating a resource that the user subscribed to has been created/updated/deleted 
    /// Several of these might come in at once
    /// </summary>
    internal class Notification
    {
        /// <summary>
        /// Gets or sets the type of change.
        /// </summary>
        [JsonProperty(PropertyName = "changeType")]
        public string ChangeType { get; set; }

        /// <summary>
        /// Gets or sets the client state used to verify that the notification is from Microsoft Graph.
        /// Compare the value received with the notification to the value you sent with the subscription request.
        /// </summary>
        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        /// <summary>
        /// Gets or sets the endpoint of the resource that changed.
        /// For example, a message uses the format ../Users/{user-id}/Messages/{message-id}
        /// </summary>
        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        /// <summary>
        /// Gets or sets the UTC date and time when the webhooks subscription expires.
        /// </summary>
        [JsonProperty(PropertyName = "subscriptionExpirationDateTime")]
        public DateTimeOffset SubscriptionExpirationDateTime { get; set; }

        /// <summary>
        /// Gets or sets the unique identifier for the webhooks subscription.
        /// </summary>
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        /// <summary>
        /// Gets or sets the properties of the changed resource.
        /// </summary>
        [JsonProperty(PropertyName = "resourceData")]
        public ResourceData ResourceData { get; set; }
    }

    /// <summary>
    /// From within a given Notification
    /// Message, Contact, and Calendar all contain this ResourceData. OneDrive does not.
    /// </summary>
    internal class ResourceData
    {
        /// <summary>
        /// Gets or sets the ID of the resource.
        /// </summary>
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the OData etag property.
        /// </summary>
        [JsonProperty(PropertyName = "@odata.etag")]
        public string ODataEtag { get; set; }

        /// <summary>
        /// Gets or sets the OData ID of the resource. This is the same value as the resource property.
        /// </summary>
        [JsonProperty(PropertyName = "@odata.id")]
        public string ODataId { get; set; }

        /// <summary>
        /// Gets or sets the OData type of the resource
        /// "#Microsoft.Graph.Message", "#Microsoft.Graph.Event", or "#Microsoft.Graph.Contact".
        /// </summary>
        [JsonProperty(PropertyName = "@odata.type")]
        public string ODataType { get; set; }
    }

    /// <summary>
    /// Helper class for change types
    /// </summary>
    public class ChangeTypeExtension
    {
        /// <summary>
        /// Convert an array of ChangeTypes to a Microsoft Graph-friendly list
        /// </summary>
        /// <param name="array">Array of change types</param>
        /// <returns>lowercase array of strings representing the change types</returns>
        public static string ConvertArrayToString(ChangeType[] array)
        {
            List<string> result = new List<string>();
            foreach (ChangeType ct in array)
            {
                result.Add(ct.ToString().ToLower());
            }

            return string.Join(", ", result);
        }
    }
}