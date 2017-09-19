// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    /// <summary>
    /// Determines what action the output binding for GraphWebhookActionAttribute
    /// will perform. Note that different modes have different semantics for the strings
    /// received by the IAsyncCollector.
    /// </summary>
    public enum GraphWebhookSubscriptionAction
    {
        /// <summary>
        /// Creates a new webhook (string = clientState)
        /// </summary>
        Create,
        /// <summary>
        /// Deletes a webhook (string = subscriptionId)
        /// </summary>
        Delete,
        /// <summary>
        /// Refreshes a webhook (string = subscriptionId)
        /// </summary>
        Refresh,
    }
}
