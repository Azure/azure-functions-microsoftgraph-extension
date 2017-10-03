// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    using Microsoft.Graph;

    internal class SubscriptionEntry
    {
        /// <summary>
        /// Gets or sets subscription ID returned by MS Graph after creation
        /// </summary>
        public Subscription Subscription { get; set; }

        /// <summary>
        /// Gets or sets the user id for the subscription
        /// </summary>
        public string UserId { get; set; } // $$$ Gets an auth token and client
    }
}
