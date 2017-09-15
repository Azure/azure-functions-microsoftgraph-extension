// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Description;
    using Microsoft.Azure.WebJobs.Extensions;
    using System;
    using TokenBinding;

    // This is filling out a Graph Subscription object. 
    // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/subscription
    [Binding]
    public class GraphWebhookAttribute : GraphTokenAttribute
    {
        /// <summary>
        /// Resource for which we're subscribing to changes.
        /// ie: me/mailFolders('Inbox')/messages
        /// </summary>
        [AutoResolve]
        public string Listen { get; set; }

        /// <summary>
        /// Gets or sets types of changes that we're subscribing to.
        /// This is specific to the resource
        /// </summary>
        public ChangeType[] ChangeTypes { get; set; }

        public GraphWebhookAction Action { get; set;}

        internal void Validate()
        {
            switch (Action)
            {
                case GraphWebhookAction.Create:
                    if (string.IsNullOrEmpty(Listen))
                    {
                        throw new ArgumentException($"A value for listen must be provided in ${Action} mode.");
                    }

                    if (ChangeTypes == null || ChangeTypes.Length == 0)
                    {
                        ChangeTypes = new ChangeType[] { ChangeType.Created, ChangeType.Deleted, ChangeType.Updated };
                    }

                    break;
                case GraphWebhookAction.Delete:
                case GraphWebhookAction.Refresh:
                    if (!string.IsNullOrEmpty(Listen))
                    {
                        throw new ArgumentException($"No value should be provided for listen in {Action} mode.");
                    }

                    if (ChangeTypes != null && ChangeTypes.Length > 0)
                    {
                        throw new ArgumentException($"No values should be provided for changeTypes in ${Action} mode.");
                    }

                    break;
            }
        }
    }
}
