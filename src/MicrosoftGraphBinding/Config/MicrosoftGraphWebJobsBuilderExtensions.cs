// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    using System;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Extensions.DependencyInjection;

    internal static class MicrosoftGraphWebJobsBuilderExtensions
    {
        public static IWebJobsBuilder AddMicrosoftGraph(this IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<MicrosoftGraphExtensionConfigProvider>()
                .Services
                .AddAuthTokenServices()
                .AddSingleton<IAsyncConverter<TokenBaseAttribute, string>, TokenConverter>()
                .AddSingleton<IGraphServiceClientProvider, GraphServiceClientProvider>()
                .AddSingleton<IGraphSubscriptionStore, WebhookSubscriptionStore>()
                .AddOptions<GraphOptions>().Configure<INameResolver>((option, appSettings) =>
                {
                    option.SetAppSettings(appSettings);
                })
                .Services
                .AddOptions<TokenOptions>().Configure<INameResolver>((option, appSettings) =>
                {
                    option.SetAppSettings(appSettings);
                });

            return builder;
        }

        public static IWebJobsBuilder AddMicrosoftGraphForTests(this IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<MicrosoftGraphExtensionConfigProvider>();
            return builder;
        }
    }
}
