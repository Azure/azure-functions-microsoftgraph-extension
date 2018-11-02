// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Token.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using Moq;

    public class TestHelpers
    {
        private static string clientSecretFile = "client_secret.txt";

        public static async Task<OutputContainer> RunTestAsync<T>(string methodName, INameResolver appSettings = null, IEasyAuthClient easyAuthClient = null, IAadClient aadClient = null,  object argument = null)
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
                    easyAuthClient = easyAuthClient ?? new Mock<IEasyAuthClient>().Object;
                    aadClient = aadClient ?? new Mock<IAadClient>().Object;
                    appSettings = appSettings ?? new Mock<INameResolver>().Object;
                    services.AddSingleton<ITypeLocator>(new FakeTypeLocator<T>());
                    services.AddSingleton<IEasyAuthClient>(easyAuthClient);
                    services.AddSingleton<IAadClient>(aadClient);
                    services.AddSingleton<INameResolver>(appSettings);
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

        public static TokenOptions GetValidSettingsForTests()
        {
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();

            if (string.IsNullOrEmpty(config[Constants.ClientSecretName]))
            {
                config[Constants.ClientSecretName] = GetClientSecret();
            }

            var resolver = new DefaultNameResolver(config);

            var options = new TokenOptions();
            options.SetAppSettings(resolver);

            return options;
        }

        public static string GetClientSecret()
        {
            if (File.Exists(clientSecretFile))
            {
                return File.ReadAllText(clientSecretFile);
            }

            return string.Empty;
        }
    }
}