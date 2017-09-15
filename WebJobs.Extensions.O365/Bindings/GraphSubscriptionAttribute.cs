// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using Microsoft.Azure.WebJobs.Description;
    using System;

    // This is used to retrieve the locally stored subscriptions
    [Binding]
    public class GraphSubscriptionAttribute : Attribute
    {   
        /// <summary>
        /// Gets or sets the UserId to filter subscriptions on; Optional
        /// </summary>
        public string UserId { get; set; }
    }
}
