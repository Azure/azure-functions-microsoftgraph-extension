// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    using Microsoft.Azure.WebJobs.Description;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph;

    public abstract class GraphTokenBaseAttribute : TokenBaseAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GraphTokenBaseAttribute"/> class.
        /// </summary>
        public GraphTokenBaseAttribute()
        {
            this.Resource = O365Constants.GraphBaseUrl;
        }
    }
}
