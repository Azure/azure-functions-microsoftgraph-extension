// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    using System;
    using Microsoft.Azure.WebJobs.Description;

    // This is used to retrieve the locally stored subscriptions
    [Binding]
    public class GraphWebhookSubscriptionAttribute : GraphTokenAttribute
    {
        private string _filter;

        /// <summary>
        /// Gets or sets the UserId to filter subscriptions on; Optional
        /// </summary>
        public string Filter
        {
            get
            {
                return _filter;
            }
            set
            {
                if (TokenIdentityMode.UserFromRequest.ToString().Equals(value, StringComparison.OrdinalIgnoreCase))
                {
                    Identity = TokenIdentityMode.UserFromRequest;   
                }
                _filter = value;
            }
        }
        
        /// <summary>
        /// Resource for which we're subscribing to changes.
        /// ie: me/mailFolders('Inbox')/messages
        /// </summary>
        [AutoResolve]
        public string SubscriptionResource { get; set; }

        /// <summary>
        /// Gets or sets types of changes that we're subscribing to.
        /// This is specific to the resource
        /// </summary>
        public GraphWebhookChangeType[] ChangeTypes { get; set; }

        public GraphWebhookSubscriptionAction Action { get; set; }

        internal void Validate()
        {
            switch (Action)
            {
                case GraphWebhookSubscriptionAction.Create:
                    if (string.IsNullOrEmpty(SubscriptionResource))
                    {
                        throw new ArgumentException($"A value for {nameof(SubscriptionResource)} must be provided in ${Action} mode.");
                    }

                    if (ChangeTypes == null || ChangeTypes.Length == 0)
                    {
                        ChangeTypes = new GraphWebhookChangeType[] { GraphWebhookChangeType.Created, GraphWebhookChangeType.Deleted, GraphWebhookChangeType.Updated };
                    }

                    break;
                case GraphWebhookSubscriptionAction.Delete:
                case GraphWebhookSubscriptionAction.Refresh:
                    if (!string.IsNullOrEmpty(SubscriptionResource))
                    {
                        throw new ArgumentException($"No value should be provided for {nameof(SubscriptionResource)} in {Action} mode.");
                    }

                    if (ChangeTypes != null && ChangeTypes.Length > 0)
                    {
                        throw new ArgumentException($"No values should be provided for {nameof(ChangeTypes)} in ${Action} mode.");
                    }

                    break;
            }
        }
    }
}
