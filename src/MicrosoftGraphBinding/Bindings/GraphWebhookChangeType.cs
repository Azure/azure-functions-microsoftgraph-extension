// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    /// <summary>
    /// Enum of changes that can be subscribed to
    /// </summary>
    public enum GraphWebhookChangeType
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
}
