// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    using System;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Graph;

    internal class GraphServiceClientProvider : IGraphServiceClientProvider
    {
        public IGraphServiceClient CreateNewGraphServiceClient(string token)
        {
            return new GraphServiceClient(
            new DelegateAuthenticationProvider(
                (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                    return Task.CompletedTask;
                }));
        }

        public void UpdateGraphServiceClientAuthToken(IGraphServiceClient client, string token)
        {
            GraphServiceClient typedClient = client as GraphServiceClient;
            if (typedClient == null)
            {
                throw new InvalidOperationException($"Only {nameof(IGraphServiceClient)} of type {nameof(GraphServiceClient)} should be created with this client provider");
            }

            typedClient.AuthenticationProvider = new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

                            return Task.CompletedTask;
                        });
        }
    }
}
