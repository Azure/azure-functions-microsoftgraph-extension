// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Extensions.Configuration.Memory;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using Moq;

    internal class TestHelpers
    {
        public static async Task<OutputContainer> RunTestAsync<T>(string methodName, IDictionary<string, string> options = null, IEasyAuthClient easyAuthClient = null, IAadServiceFactory aadServiceFactory = null, HttpRequest request = null)
        {
            var outputContainer = new OutputContainer();
            var arguments = new Dictionary<string, object>()
            {
                { "outputContainer", outputContainer },
            };

            if (request != null)
            {
                arguments.Add("$request", request);
            }

            IHost host = new HostBuilder()
                .ConfigureAppConfiguration(c=> {

                    if (options != null)
                    {
                        c.Sources.Clear();

                        var source = new MemoryConfigurationSource
                        {
                            InitialData = options
                        };

                        c.Add(source);
                    }
                })
                .ConfigureServices(services =>
                {
                    easyAuthClient = easyAuthClient ?? new Mock<IEasyAuthClient>().Object;
                    aadServiceFactory = aadServiceFactory ?? new Mock<IAadServiceFactory>().Object;
                    services.AddSingleton<ITypeLocator>(new FakeTypeLocator<T>());
                    services.AddSingleton<IEasyAuthClient>(easyAuthClient);
                    services.AddSingleton<IAadServiceFactory>(aadServiceFactory);
                })
                .ConfigureWebJobs(builder =>
                {
                    builder.AddAuthTokenForTests();
                    builder.UseHostId(Guid.NewGuid().ToString("n"));
                })
                .Build();

            JobHost webJobsHost = host.Services.GetService<IJobHost>() as JobHost;
            await webJobsHost.CallAsync(methodName, arguments);
            return outputContainer;
        }
    }
}