// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using System;
    using Microsoft.Azure.WebJobs.Description;

    [Binding]
    public class GraphWebhookTriggerAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets oData type of payload
        /// "#Microsoft.Graph.Message", "#Microsoft.Graph.DriveItem", "#Microsoft.Graph.Contact", "#Microsoft.Graph.Event"
        /// </summary>
        public string ResourceType { get; set; }
    }
}