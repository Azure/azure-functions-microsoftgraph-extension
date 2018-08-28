// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;
using Microsoft.Extensions.DependencyInjection;

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    public static class AuthTokenWebJobsBuilderExtensions
    {
        public static IWebJobsBuilder AddAuthToken(this IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<AuthTokenExtensionConfigProvider>()
                .BindOptions<TokenOptions>()
                .Services
                .AddSingleton<IEasyAuthClient, EasyAuthTokenClient>()
                .AddSingleton<IAadClient, AadClient>();
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
                .AddSingleton<IAadClient, AadClient>();
            return services;
        }
    }
}
