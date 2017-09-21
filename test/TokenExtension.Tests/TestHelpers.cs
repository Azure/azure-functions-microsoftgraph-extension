// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using System;

    public class TestHelpers
    {
        public static JobHost NewHost<T>(AuthTokenExtensionConfig ext = null)
        {
            JobHostConfiguration config = new JobHostConfiguration();
            config.HostId = Guid.NewGuid().ToString("n");
            config.StorageConnectionString = null;
            config.DashboardConnectionString = null;
            config.TypeLocator = new FakeTypeLocator<T>();
            config.AddExtension(ext ?? new AuthTokenExtensionConfig());
            var host = new JobHost(config);
            return host;
        }
    }
}