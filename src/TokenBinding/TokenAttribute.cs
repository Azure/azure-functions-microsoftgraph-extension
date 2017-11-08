// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using Microsoft.Azure.WebJobs.Description;

    //This class exists to allow rules to explicitly bind to the token attribute, while still having the logic
    //contained in an abstract base class so that the Graph extension attributes can extend from that class. This prevents
    //conflicts from an ExcelAttribute being both an ExcelAttribute and a TokenAttribute when determining what rules to follow.
    [Binding]
    public sealed class TokenAttribute : TokenBaseAttribute
    {
    }
}
