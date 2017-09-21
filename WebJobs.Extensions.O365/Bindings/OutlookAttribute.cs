// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    using Microsoft.Azure.WebJobs.Description;

    /// <summary>
    /// Outlook Attribute inherits from TokenAttribute
    /// No additional info necessary, but remains separate class in order to maintain clarity
    /// </summary>
    [Binding]
    public class OutlookAttribute : GraphTokenAttribute
    {
    }
}
