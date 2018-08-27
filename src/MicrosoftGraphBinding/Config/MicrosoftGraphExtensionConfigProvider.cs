// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests")]
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]
namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.IO;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;
    using static Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters.ExcelConverters;
    using static Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters.GraphWebhookSubscriptionConverters;

    /// <summary>
    /// WebJobs SDK Extension for O365 Token binding.
    /// </summary>
    internal class MicrosoftGraphExtensionConfigProvider : IExtensionConfigProvider,
        IAsyncConverter<HttpRequestMessage, HttpResponseMessage>
    {
        private readonly GraphServiceClientManager _graphServiceClientManager;
        private readonly IGraphSubscriptionStore _subscriptionStore;
        private readonly ILoggerFactory _loggerFactory;
        private readonly GraphOptions _options;
        private Uri _notificationUrl;
        private WebhookTriggerBindingProvider _webhookTriggerProvider;

        public MicrosoftGraphExtensionConfigProvider(IOptions<GraphOptions> options, 
            ILoggerFactory loggerFactory, 
            IGraphServiceClientProvider graphClientProvider, 
            INameResolver appSettings,
            IAsyncConverter<TokenAttribute, string> tokenConverter,
            IGraphSubscriptionStore subscriptionStore)
        {
            _options = options.Value;
            _options.SetAppSettings(appSettings);
            _graphServiceClientManager = new GraphServiceClientManager(_options, tokenConverter, graphClientProvider);
            _subscriptionStore = subscriptionStore;
            _loggerFactory = loggerFactory;
        }

        /// <summary>
        /// Initialize the O365 binding extension
        /// </summary>
        /// <param name="context">Context containing info relevant to this extension</param>
        public void Initialize(ExtensionConfigContext context)
        {
            _webhookTriggerProvider = new WebhookTriggerBindingProvider();
            _notificationUrl = context.GetWebhookHandler();

            var graphWebhookConverter = new GraphWebhookSubscriptionConverter(_graphServiceClientManager, _options, _subscriptionStore);

            // Webhooks
            var webhookSubscriptionRule = context.AddBindingRule<GraphWebhookSubscriptionAttribute>();
            webhookSubscriptionRule.BindToInput<Subscription[]>(graphWebhookConverter);
            webhookSubscriptionRule.BindToInput<OpenType[]>(typeof(GenericGraphWebhookSubscriptionConverter<>), _graphServiceClientManager, _options, _subscriptionStore);
            webhookSubscriptionRule.BindToInput<string[]>(graphWebhookConverter);
            webhookSubscriptionRule.BindToInput<JArray>(graphWebhookConverter);
            webhookSubscriptionRule.BindToCollector<string>(CreateCollector);
            context.AddBindingRule<GraphWebhookTriggerAttribute>().BindToTrigger(_webhookTriggerProvider);

            // OneDrive
            var oneDriveService = new OneDriveService(_graphServiceClientManager);
            var OneDriveRule = context.AddBindingRule<OneDriveAttribute>();
            var oneDriveConverter = new OneDriveConverter(oneDriveService);

            // OneDrive inputs
            OneDriveRule.BindToInput<byte[]>(oneDriveConverter);
            OneDriveRule.BindToInput<string>(oneDriveConverter);
            OneDriveRule.BindToInput<Stream>(oneDriveConverter);
            OneDriveRule.BindToInput<DriveItem>(oneDriveConverter);
            //OneDriveoutputs
            OneDriveRule.BindToCollector<byte[]>(oneDriveConverter);

            // Excel
            var excelService = new ExcelService(_graphServiceClientManager);
            var ExcelRule = context.AddBindingRule<ExcelAttribute>();
            var excelConverter = new ExcelConverter(excelService);
            // Excel Outputs
            ExcelRule.AddConverter<object[][], string>(ExcelService.CreateRows);
            ExcelRule.AddConverter<JObject, string>(excelConverter);
            ExcelRule.AddOpenConverter<OpenType, string>(typeof(ExcelGenericsConverter<>), excelService); // used to append/update arrays of POCOs
            ExcelRule.BindToCollector<string>(excelConverter);
            // Excel Inputs
            ExcelRule.BindToInput<string[][]>(excelConverter);
            ExcelRule.BindToInput<WorkbookTable>(excelConverter);
            ExcelRule.BindToInput<OpenType>(typeof(ExcelGenericsConverter<>), excelService);

            // Outlook
            var outlookService = new OutlookService(_graphServiceClientManager);
            var OutlookRule = context.AddBindingRule<OutlookAttribute>();
            var outlookConverter = new OutlookConverter(outlookService);
            // Outlook Outputs           
            OutlookRule.AddConverter<JObject, Message>(outlookConverter);
            OutlookRule.AddOpenConverter<OpenType, Message>(typeof(OutlookGenericsConverter<>), outlookService);
            OutlookRule.AddConverter<string, Message>(outlookConverter);
            OutlookRule.BindToCollector<Message>(outlookConverter);
        }

        private IAsyncCollector<string> CreateCollector(GraphWebhookSubscriptionAttribute attr)
        {
            return new GraphWebhookSubscriptionAsyncCollector(_graphServiceClientManager, _options, _loggerFactory, _subscriptionStore, _notificationUrl, attr);
        }

        //TODO: https://github.com/Azure/azure-functions-microsoftgraph-extension/issues/48
        internal static string CreateBindingCategory(string bindingName)
        {
            return $"Host.Bindings.{bindingName}";
        }

        /// <summary>
        /// HttpRequest -> HttpResponse
        /// Used to create new MSGraph subscriptions
        /// </summary>
        /// <param name="input">HttpRequestMessage from fx</param>
        /// <param name="cancellationToken">CancellationToken cancellationToken</param>
        /// <returns>Task with HttpResponseMessage for further processing</returns>
        async Task<HttpResponseMessage> IAsyncConverter<HttpRequestMessage, HttpResponseMessage>.ConvertAsync(HttpRequestMessage input, CancellationToken cancellationToken)
        {
            var handler = new GraphWebhookSubscriptionHandler(_graphServiceClientManager, _subscriptionStore, _loggerFactory, _notificationUrl, _webhookTriggerProvider);
            var response = await handler.ProcessAsync(input, cancellationToken);
            return response;
        }
    }
}
