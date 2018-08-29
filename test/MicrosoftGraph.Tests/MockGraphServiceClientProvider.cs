// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Graph;

    class MockGraphServiceClientProvider : IGraphServiceClientProvider
    {
        private IGraphServiceClient _client;

        public MockGraphServiceClientProvider(IGraphServiceClient client)
        {
            _client = client;
        }

        public IGraphServiceClient CreateNewGraphServiceClient(string token)
        {
            return _client;
        }


        public void UpdateGraphServiceClientAuthToken(IGraphServiceClient client, string token)
        {
            //NO-OP
        }
    }
}
