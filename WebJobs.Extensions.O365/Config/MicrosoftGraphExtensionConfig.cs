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
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Host;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// WebJobs SDK Extension for O365 Token binding.
    /// </summary>
    public class MicrosoftGraphExtensionConfig : IExtensionConfigProvider,
        IAsyncConverter<HttpRequestMessage, HttpResponseMessage>
    {
        internal IExcelClient _excelClient { get; set; }

        internal IOutlookClient _outlookClient { get; set; }

        internal IOneDriveClient _onedriveClient { get; set; }

        internal ServiceManager _serviceManager { get; set; }

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
            this._appSettings = config.NameResolver;

            // Set up logging
            _log = context.Trace;

            ConfigureServiceManager(context);

            // Infer a blank Notification URL from the appsettings.
            string appSettingBYOBTokenMap = _appSettings.Resolve(O365Constants.AppSettingBYOBTokenMap);
            var subscriptionStore = new WebhookSubscriptionStore(appSettingBYOBTokenMap);
            var webhookTriggerProvider = new WebhookTriggerBindingProvider();
            _webhookConfig = new GraphWebhookConfig(context.GetWebhookHandler(), subscriptionStore, webhookTriggerProvider);

            var converter = new Converters(_serviceManager, _webhookConfig);

            // Extend token attribute to retrieve [authenticated] GraphServiceClient
            //this.tokenExtension.TokenRule.BindToInput<GraphServiceClient>(converter);

            // Webhooks
            var webhookSubscriptionRule = context.AddBindingRule<GraphWebhookSubscriptionAttribute>();

            webhookSubscriptionRule.BindToInput<Subscription[]>(converter);
            webhookSubscriptionRule.BindToInput<string[]>(converter);
            webhookSubscriptionRule.BindToCollector<string>(CreateCollector);

            context.AddBindingRule<GraphWebhookTriggerAttribute>().BindToTrigger(webhookTriggerProvider);

            // OneDrive
            var OneDriveRule = context.AddBindingRule<OneDriveAttribute>();

            // OneDrive inputs
            OneDriveRule.BindToInput<byte[]>(converter);
            OneDriveRule.BindToInput<string>(converter);
            OneDriveRule.BindToInput<Stream>(converter);
            OneDriveRule.BindToInput<DriveItem>(converter);

            // OneDrive Outputs
            OneDriveRule.AddConverter<byte[], Stream>(OneDriveService.CreateStream);
            OneDriveRule.BindToCollector<Stream>(converter.CreateCollector);

            // Excel
            var ExcelRule = context.AddBindingRule<ExcelAttribute>();

            // Excel Outputs
            ExcelRule.AddConverter<object[][], JObject>(ExcelService.CreateRows);
            ExcelRule.AddConverter<List<OpenType>, JObject>(typeof(GenericConverter<>)); // used to append/update lists of POCOs
            ExcelRule.AddConverter<OpenType, JObject>(typeof(GenericConverter<>)); // used to append/update arrays of POCOs
            ExcelRule.BindToCollector<JObject>(converter.CreateCollector);
            ExcelRule.BindToCollector<JObject>(typeof(POCOConverter<>));

            // Excel Inputs
            ExcelRule.BindToInput<string[][]>(converter);
            ExcelRule.BindToInput<WorkbookTable>(converter);
            ExcelRule.BindToInput<List<OpenType>>(typeof(POCOConverter<>), _serviceManager);
            ExcelRule.BindToInput<OpenType>(typeof(POCOConverter<>), _serviceManager);

            // Outlook
            var OutlookRule = context.AddBindingRule<OutlookAttribute>();

            // Outlook Outputs
            OutlookRule.AddConverter<JObject, Message>(OutlookService.CreateMessage);
            OutlookRule.AddConverter<string, Message>(OutlookService.CreateMessage);
            OutlookRule.BindToCollector<Message>(converter.CreateCollector);
        }

        private void ConfigureServiceManager(ExtensionConfigContext context)
        {
            // Set up token extension; handles auth (only providers supported by Easy Auth)
            var tokenExtension = new AuthTokenExtensionConfig();
            tokenExtension.InitializeAllExceptRules(context);
            _serviceManager = new ServiceManager(tokenExtension);
            _serviceManager.ExcelClient = _excelClient;
            _serviceManager.OutlookClient = _outlookClient;
            _serviceManager.OneDriveClient = _onedriveClient;
        }

        public IAsyncCollector<string> CreateCollector(GraphWebhookSubscriptionAttribute attr)
        {
            return new GraphWebhookSubscriptionAsyncCollector(_serviceManager, _log, _webhookConfig, attr);
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

        /// <summary>
        /// Used to convert POCOs to JObjects (for Excel output bindings)
        /// T -> used to append a row
        /// T[] -> used to update a table
        /// </summary>
        /// <typeparam name="T">Generic POCO type</typeparam>
        public class GenericConverter<T> : IConverter<List<T>, JObject>, IConverter<T, JObject>
        {
            /// <summary>
            /// Convert from POCO -> JObject (either row or rows)
            /// </summary>
            /// <param name="input">POCO input from fx</param>
            /// <returns>JObject with proper keys set</returns>
            public JObject Convert(T input)
            {
                // handle T[]
                if (typeof(T).IsArray)
                {
                    var array = input as object[];
                    return ConvertEnumerable(array);
                }
                else
                {
                    // handle T
                    JObject data = JObject.FromObject(input);
                    data[O365Constants.POCOKey] = true; // Set Microsoft.O365Bindings.POCO flag to indicate that data is from POCO (vs. object[][])

                    return data;
                }
            }

            /// <summary>
            /// Convert from List<POCO> -> JObject
            /// </summary>
            /// <param name="input">POCO input from fx</param>
            /// <returns>JObject with proper keys set</returns>
            public JObject Convert(List<T> input)
            {
                return ConvertEnumerable(input);
            }

            private JObject ConvertEnumerable<U>(IEnumerable<U> input)
            {
                JObject jsonContent = new JObject();

                JArray rowData = JArray.FromObject(input);

                // List<T> -> JArray
                jsonContent[O365Constants.ValuesKey] = rowData;

                // Set rows, columns needed if updating entire worksheet
                jsonContent[O365Constants.RowsKey] = rowData.Count();

                // No exception -- array is rectangular by default
                jsonContent[O365Constants.ColsKey] = rowData.First.Count();

                // Set POCO key to indicate that the values need to be ordered to match the header of the existing table
                jsonContent[O365Constants.POCOKey] = true;

                return jsonContent;
            }
        }

        /// <summary>
        /// Used for INPUT bindings: convert Excel Attribute -> POCO inputs
        /// </summary>
        /// <typeparam name="T">POCO type user wishes to bind Excel contents to</typeparam>
        internal class POCOConverter<T> : IAsyncConverter<ExcelAttribute, T[]>, IAsyncConverter<ExcelAttribute, List<T>>
            where T : new()
        {
            private readonly ServiceManager parent;

            /// <summary>
            /// Initializes a new instance of the <see cref="POCOConverter{T}"/> class.
            /// </summary>
            /// <param name="parent">O365Extension to which the result of the request for data will be returned</param>
            public POCOConverter(ServiceManager parent)
            {
                this.parent = parent;
            }

            async Task<List<T>> IAsyncConverter<ExcelAttribute, List<T>>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var manager = this.parent.GetExcelManager(input);
                var result = await manager.GetExcelRangePOCOListAsync<T>(input);
                return result;
            }

            async Task<T[]> IAsyncConverter<ExcelAttribute, T[]>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var manager = this.parent.GetExcelManager(input);
                var result = manager.GetExcelRangePOCOAsync<T>(input);
                return await result;
            }

            public IAsyncCollector<JObject> CreateCollector(ExcelAttribute attr)
            {
                var manager = this.parent.GetExcelManager(attr);
                return new ExcelAsyncCollector(manager, attr);
            }
        }

        /// <summary>
        /// Used for input bindings; Attribute -> Input type
        /// </summary>
        internal class Converters :
            IAsyncConverter<ExcelAttribute, string[][]>,
            IAsyncConverter<ExcelAttribute, WorkbookTable>,
            IAsyncConverter<OneDriveAttribute, byte[]>,
            IAsyncConverter<OneDriveAttribute, string>,
            IAsyncConverter<OneDriveAttribute, Stream>,
            IAsyncConverter<OneDriveAttribute, DriveItem>,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, Subscription[]>,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, string[]>
        {
            private readonly ServiceManager _serviceManager;
            private readonly GraphWebhookConfig _webhookConfig;

            public Converters(ServiceManager parent, GraphWebhookConfig webhookConfig)
            {
                _serviceManager = parent;
                _webhookConfig = webhookConfig;
            }

            public IAsyncCollector<JObject> CreateCollector(ExcelAttribute attr)
            {
                var service = _serviceManager.GetExcelManager(attr);
                return new ExcelAsyncCollector(service, attr);
            }

            public IAsyncCollector<Stream> CreateCollector(OneDriveAttribute attr)
            {
                var service = _serviceManager.GetOneDriveService(attr);
                return new OneDriveAsyncCollector(service, attr);
            }

            public IAsyncCollector<Message> CreateCollector(OutlookAttribute attr)
            {
                return new OutlookAsyncCollector(_serviceManager.GetOutlookService(attr));
            }



            async Task<string[][]> IAsyncConverter<ExcelAttribute, string[][]>.ConvertAsync(ExcelAttribute attr, CancellationToken cancellationToken)
            {
                var service = _serviceManager.GetExcelManager(attr);
                var result = await service.GetExcelRangeAsync(attr);
                return result;
            }

            async Task<WorkbookTable> IAsyncConverter<ExcelAttribute, WorkbookTable>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var service = _serviceManager.GetExcelManager(input);
                var result = await service.GetExcelTable(input);
                return result;
            }

            async Task<byte[]> IAsyncConverter<OneDriveAttribute, byte[]>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            { 
                var service = _serviceManager.GetOneDriveService(input);

                var result = await service.GetOneDriveContentsAsByteArrayAsync(input);

                return result;
            }

            async Task<string> IAsyncConverter<OneDriveAttribute, string>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var service = _serviceManager.GetOneDriveService(input);

                var byteArray = await service.GetOneDriveContentsAsByteArrayAsync(input);

                return Encoding.UTF8.GetString(byteArray);
            }

            async Task<Stream> IAsyncConverter<OneDriveAttribute, Stream>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var service = _serviceManager.GetOneDriveService(input);

                return await service.GetOneDriveContentsAsStreamAsync(input);
            }

            async Task<DriveItem> IAsyncConverter<OneDriveAttribute, DriveItem>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var service = _serviceManager.GetOneDriveService(input);
                return await service.GetOneDriveItemAsync(input);
            }

            async Task<Subscription[]> IAsyncConverter<GraphWebhookSubscriptionAttribute, Subscription[]>.ConvertAsync(GraphWebhookSubscriptionAttribute input, CancellationToken cancellationToken)
            {
                return await GetSubscriptionsFromAttribute(input);
            }

            async Task<string[]> IAsyncConverter<GraphWebhookSubscriptionAttribute, string[]>.ConvertAsync(GraphWebhookSubscriptionAttribute input, CancellationToken cancellationToken)
            {
                Subscription[] subscriptions = await GetSubscriptionsFromAttribute(input);
                return subscriptions.Select(sub => sub.Id).ToArray();
            }

            private async Task<Subscription[]> GetSubscriptionsFromAttribute(GraphWebhookSubscriptionAttribute attribute)
            {
                IEnumerable<WebhookSubscriptionStore.SubscriptionEntry> subscriptionEntries = await _webhookConfig.SubscriptionStore.GetAllSubscriptionsAsync();
                if (TokenIdentityMode.UserFromRequest.ToString().Equals(attribute.Filter, StringComparison.OrdinalIgnoreCase))
                {
                    var dummyTokenAttribute = new TokenAttribute()
                    {
                        Resource = O365Constants.GraphBaseUrl,
                        Identity = TokenIdentityMode.UserFromToken,
                        UserToken = attribute.UserToken,
                        IdentityProvider = "AAD",
                    };
                    var graph = await _serviceManager.GetMSGraphClientAsync(dummyTokenAttribute);
                    var user = await graph.Me.Request().GetAsync();
                    subscriptionEntries = subscriptionEntries.Where(entry => entry.UserId.Equals(user.Id));
                }
                else if (attribute.Filter != null)
                {
                    throw new InvalidOperationException($"There is no filter for {attribute.Filter}");
                }

                return subscriptionEntries.Select(entry => entry.Subscription).ToArray();
            }
        }
    }
}
