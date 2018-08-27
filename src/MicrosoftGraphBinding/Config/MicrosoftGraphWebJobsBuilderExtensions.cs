using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.DependencyInjection;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    internal static class MicrosoftGraphWebJobsBuilderExtensions
    {
        public static IWebJobsBuilder AddMicrosoftGraph(this IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<MicrosoftGraphExtensionConfigProvider>()
                .BindOptions<GraphOptions>()
                .Services
                .AddSingleton<IGraphServiceClientProvider, GraphServiceClientProvider>()
                .AddSingleton<IGraphSubscriptionStore, IGraphSubscriptionStore>();
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
