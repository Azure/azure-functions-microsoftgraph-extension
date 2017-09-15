// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Azure.WebJobs;

    public class FakeTypeLocator<T> : ITypeLocator
    {
        public IReadOnlyList<Type> GetTypes()
        {
            return new Type[] { typeof(T) };
        }
    }
}