// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    class Mock : ServiceManager
    {
        private IGraphServiceClient _client;

        public Mock(AuthTokenExtensionConfig config, IGraphServiceClient client) : base(config)
        {
            _client = client;
        }

        public override Task<IGraphServiceClient> GetMSGraphClientAsync(TokenBaseAttribute attribute)
        {
            return Task.FromResult(_client);
        }
    }
}
