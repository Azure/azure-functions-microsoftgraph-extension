// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Extensions.Token.Tests;
    using Microsoft.Graph;
    using Moq;

    internal static class CommonUtilities
    {
        public static async Task ExecuteFunction<T>(Mock<IGraphServiceClient> clientMock, string functionReference)
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            ServiceManager manager = new Mock(null, clientMock.Object);
            graphConfig._serviceManager = manager;

            var jobHost = TestHelpers.NewHost<T>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync(functionReference, args);
        }

    }
}
