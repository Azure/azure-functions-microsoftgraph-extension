// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Azure.WebJobs.Extensions.Token.Tests;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Graph;
    using Moq;

    internal static class CommonUtilities
    {
        public static async Task<OutputContainer> ExecuteFunction<T>(string methodName, Mock<IGraphServiceClient> client, IGraphSubscriptionStore subscriptionStore = null, INameResolver appSettings = null, object argument = null)
        {
            var outputContainer = new OutputContainer();
            var arguments = new Dictionary<string, object>()
            {
                { "outputContainer", outputContainer },
                { "triggerData", argument }
            };

            IHost host = new HostBuilder()
                .ConfigureServices(services =>
                {
                    appSettings = appSettings ?? new DefaultNameResolver();
                    subscriptionStore = subscriptionStore ?? new MemorySubscriptionStore();
                    services.AddSingleton<ITypeLocator>(new FakeTypeLocator<T>());
                    services.AddSingleton<IAsyncConverter<TokenBaseAttribute, string>>(new MockTokenConverter());
                    services.AddSingleton<IGraphServiceClientProvider>(new MockGraphServiceClientProvider(client.Object));
                    services.AddSingleton<IGraphSubscriptionStore>(subscriptionStore);
                    services.AddSingleton<INameResolver>(appSettings);
                })
                .ConfigureWebJobs(builder =>
                {
                    builder.AddMicrosoftGraphForTests();
                    builder.UseHostId(Guid.NewGuid().ToString("n"));
                })
                .Build();

            JobHost webJobsHost = host.Services.GetService<IJobHost>() as JobHost;
            await webJobsHost.CallAsync(methodName, arguments);
            return outputContainer;
        }
    }
}