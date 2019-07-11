// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{

    using System;
    using Microsoft.Extensions.DependencyInjection;

    public static class AuthTokenWebJobsBuilderExtensions
    {
        public static IWebJobsBuilder AddAuthToken(this IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<AuthTokenExtensionConfigProvider>()
                .ConfigureOptions<TokenOptions>((rootConfig, extensionPath, options) =>
                {
                    options.HostName = rootConfig[Constants.WebsiteHostname];
                    options.ClientId = rootConfig[Constants.ClientIdName];
                    options.ClientSecret = rootConfig[Constants.ClientSecretName];
                    options.TenantUrl = rootConfig[Constants.WebsiteAuthOpenIdIssuer];
                    options.SigningKey = rootConfig[Constants.WebsiteAuthSigningKey];
                })
                .Services
                .AddSingleton<IEasyAuthClient, EasyAuthTokenClient>()
                .AddSingleton<IAadServiceFactory, AdalAadServiceFactory>();
            return builder;
        }

        public static IWebJobsBuilder AddAuthTokenForTests(this IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<AuthTokenExtensionConfigProvider>();
            return builder;
        }

        public static IServiceCollection AddAuthTokenServices(this IServiceCollection services)
        {
            services.AddSingleton<IEasyAuthClient, EasyAuthTokenClient>()
                .AddSingleton<IAadService, AdalAadService>();
            return services;
        }
    }
}
