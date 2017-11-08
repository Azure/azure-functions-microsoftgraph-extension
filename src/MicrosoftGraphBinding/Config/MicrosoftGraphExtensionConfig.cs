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
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Host;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Config;
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
        internal TraceWriter _log;

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
            _log = context.Trace;

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
            ExcelRule.AddConverter<object[][], JObject>(ExcelService.CreateRows);
            ExcelRule.AddConverter<List<OpenType>, JObject>(typeof(GenericExcelRowConverter<>)); // used to append/update lists of POCOs
            ExcelRule.AddConverter<OpenType, JObject>(typeof(GenericExcelRowConverter<>)); // used to append/update arrays of POCOs
            ExcelRule.BindToCollector<JObject>(excelConverter.CreateCollector);
            ExcelRule.BindToCollector<JObject>(typeof(POCOExcelRowConverter<>));

            // Excel Inputs
            ExcelRule.BindToInput<string[][]>(excelConverter);
            ExcelRule.BindToInput<WorkbookTable>(excelConverter);
            ExcelRule.BindToInput<List<OpenType>>(typeof(POCOExcelRowConverter<>), _serviceManager);
            ExcelRule.BindToInput<OpenType>(typeof(POCOExcelRowConverter<>), _serviceManager);

            // Outlook
            var OutlookRule = context.AddBindingRule<OutlookAttribute>();

            // Outlook Outputs
            OutlookRule.AddConverter<JObject, Message>(OutlookService.CreateMessage);
            OutlookRule.AddConverter<string, Message>(OutlookService.CreateMessage);
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
            return new GraphWebhookSubscriptionAsyncCollector(_serviceManager, _log, _webhookConfig, attr);
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


        /// <summary>
        /// HttpRequest -> HttpResponse
        /// Used to create new MSGraph subscriptions
        /// </summary>
        /// <param name="input">HttpRequestMessage from fx</param>
        /// <param name="cancellationToken">CancellationToken cancellationToken</param>
        /// <returns>Task with HttpResponseMessage for further processing</returns>
        async Task<HttpResponseMessage> IAsyncConverter<HttpRequestMessage, HttpResponseMessage>.ConvertAsync(HttpRequestMessage input, CancellationToken cancellationToken)
        {
            var handler = new GraphWebhookSubscriptionHandler(_serviceManager, _webhookConfig, _log);
            var response = await handler.ProcessAsync(input);
            return response;
        }
    }
}
