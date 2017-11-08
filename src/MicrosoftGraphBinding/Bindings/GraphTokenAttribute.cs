// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    using Microsoft.Azure.WebJobs.Description;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph;

    /// <summary>
    /// Abstract attribute to be base class for all Graph-based binding attributes
    /// </summary>
    [Binding]
    public abstract class GraphTokenAttribute : TokenBaseAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GraphTokenAttribute"/> class.
        /// </summary>
        public GraphTokenAttribute()
        {
            this.Resource = O365Constants.GraphBaseUrl;
        }
    }
}
