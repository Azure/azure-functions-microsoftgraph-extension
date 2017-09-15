// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// internal trigger payload for when a graph webhook fires.
    /// </summary>
    internal class WebhookTriggerData
    {
        /// <summary>
        /// Authenticated client that can call & access the resource;
        /// should be the same client that subscribed to the notification.
        /// Used to allow other O365 bindings to get on-behalf credentials
        /// </summary>
        public GraphServiceClient client;

        /// <summary>
        /// Results of GET request made using the resource in the notification
        /// Ultimately what ends up being passed to the user's fx
        /// </summary>
        public JObject Payload;

        /// <summary>
        /// Type used for filtering notifications
        /// (Should also be in Payload)
        /// </summary>
        public string odataType;
    }
}