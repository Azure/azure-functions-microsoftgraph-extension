// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests")]
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]
namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;
    using static Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters.ExcelConverters;
    using static Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters.GraphWebhookSubscriptionConverters;

    /// <summary>
    /// WebJobs SDK Extension for O365 Token binding.
    /// </summary>
    public class MicrosoftGraphExtensionConfig : IExtensionConfigProvider,
        IAsyncConverter<HttpRequestMessage, HttpResponseMessage>
    {
        internal ServiceManager _serviceManager { get; set; }

        internal IGraphSubscriptionStore _subscriptionStore { get; set; }

        internal GraphWebhookConfig _webhookConfig;

        /// <summary>
        /// Used to confer information, warnings, etc. to function app log
        /// </summary>
        internal ILoggerFactory _loggerFactory;

        internal INameResolver _appSettings;

        /// <summary>
        /// Initialize the O365 binding extension
        /// </summary>
        /// <param name="context">Context containing info relevant to this extension</param>
        public void Initialize(ExtensionConfigContext context)
        {
            var config = context.Config;
            _appSettings = config.NameResolver;

            // Set up logging
            _loggerFactory = context.Config.LoggerFactory ?? throw new ArgumentNullException("No logger present");

            ConfigureServiceManager(context);

            // Infer a blank Notification URL from the appsettings.
            string appSettingBYOBTokenMap = _appSettings.Resolve(O365Constants.AppSettingBYOBTokenMap);
            var subscriptionStore = _subscriptionStore ?? new WebhookSubscriptionStore(appSettingBYOBTokenMap);
            var webhookTriggerProvider = new WebhookTriggerBindingProvider();
            _webhookConfig = new GraphWebhookConfig(context.GetWebhookHandler(), subscriptionStore, webhookTriggerProvider);

            var graphWebhookConverter = new GraphWebhookSubscriptionConverter(_serviceManager, _webhookConfig);

            // Webhooks
            var webhookSubscriptionRule = context.AddBindingRule<GraphWebhookSubscriptionAttribute>();
            webhookSubscriptionRule.BindToInput<Subscription[]>(graphWebhookConverter);
            webhookSubscriptionRule.BindToInput<OpenType[]>(typeof(GenericGraphWebhookSubscriptionConverter<>), _serviceManager, _webhookConfig);
            webhookSubscriptionRule.BindToInput<string[]>(graphWebhookConverter);
            webhookSubscriptionRule.BindToInput<JArray>(graphWebhookConverter);
            webhookSubscriptionRule.BindToCollector<string>(CreateCollector);

            context.AddBindingRule<GraphWebhookTriggerAttribute>().BindToTrigger(webhookTriggerProvider);

            // OneDrive
            var OneDriveRule = context.AddBindingRule<OneDriveAttribute>();
            var oneDriveConverter = new OneDriveConverter(_serviceManager);

            // OneDrive inputs
            OneDriveRule.BindToInput<byte[]>(oneDriveConverter);
            OneDriveRule.BindToInput<string>(oneDriveConverter);
            OneDriveRule.BindToInput<Stream>(oneDriveConverter);
            OneDriveRule.BindToInput<DriveItem>(oneDriveConverter);

            OneDriveRule.BindToCollector<byte[]>(CreateCollector);

            // Excel
            var ExcelRule = context.AddBindingRule<ExcelAttribute>();
            var excelConverter = new ExcelConverter(_serviceManager);

            // Excel Outputs
            ExcelRule.AddConverter<object[][], string>(ExcelService.CreateRows);
            ExcelRule.AddOpenConverter<OpenType, string>(typeof(ExcelGenericsConverter<>), _serviceManager); // used to append/update arrays of POCOs
            ExcelRule.BindToCollector<string>(excelConverter);

            // Excel Inputs
            ExcelRule.BindToInput<string[][]>(excelConverter);
            ExcelRule.BindToInput<WorkbookTable>(excelConverter);
            ExcelRule.BindToInput<OpenType>(typeof(ExcelGenericsConverter<>), _serviceManager);

            // Outlook
            var OutlookRule = context.AddBindingRule<OutlookAttribute>();
            var outlookConverter = new OutlookConverter();

            // Outlook Outputs           
            OutlookRule.AddConverter<JObject, Message>(outlookConverter);
            OutlookRule.AddOpenConverter<OpenType, Message>(typeof(OutlookGenericsConverter<>));
            OutlookRule.AddConverter<string, Message>(outlookConverter);
            OutlookRule.BindToCollector<Message>(CreateCollector);
        }

        private void ConfigureServiceManager(ExtensionConfigContext context)
        {
            if(_serviceManager == null)
            {
                // Set up token extension; handles auth (only providers supported by Easy Auth)
                var tokenExtension = new AuthTokenExtensionConfig();
                tokenExtension.InitializeAllExceptRules(context);
                _serviceManager = new ServiceManager(tokenExtension);
            }
        }

        private IAsyncCollector<string> CreateCollector(GraphWebhookSubscriptionAttribute attr)
        {
            return new GraphWebhookSubscriptionAsyncCollector(_serviceManager, _loggerFactory, _webhookConfig, attr);
        }

        private IAsyncCollector<Message> CreateCollector(OutlookAttribute attr)
        {
            var service = Task.Run(() => _serviceManager.GetOutlookService(attr)).GetAwaiter().GetResult();
            return new OutlookAsyncCollector(service);
        }

        private IAsyncCollector<byte[]> CreateCollector(OneDriveAttribute attr)
        {
            var service = Task.Run(() => _serviceManager.GetOneDriveService(attr)).GetAwaiter().GetResult();
            return new OneDriveAsyncCollector(service, attr);
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
            var handler = new GraphWebhookSubscriptionHandler(_serviceManager, _webhookConfig, _loggerFactory);
            var response = await handler.ProcessAsync(input);
            return response;
        }
    }
}
