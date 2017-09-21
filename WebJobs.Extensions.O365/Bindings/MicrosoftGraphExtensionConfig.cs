// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Azure.WebJobs.Host;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// WebJobs SDK Extension for O365 Token binding.
    /// </summary>
    public class MicrosoftGraphExtensionConfig : IExtensionConfigProvider,
        IAsyncConverter<HttpRequestMessage, HttpResponseMessage>
    {
        /// <summary>
        /// Map principal Id + scopes -> GraphServiceClient + token expiration date
        /// </summary>
        private ConcurrentDictionary<string, CachedClient> clients = new ConcurrentDictionary<string, CachedClient>();

        private AuthTokenExtensionConfig tokenExtension;

        // Cache for communicating tokens across webhooks
        internal WebhookSubscriptionStore subscriptionStore;

        private WebhookTriggerBindingProvider webhookTriggerProvider;

        /// <summary>
        /// Used to confer information, warnings, etc. to function app log
        /// </summary>
        internal TraceWriter Log;

        /// <summary>
        /// Gets or sets optional Url if we allow subscribing for WebHooks. Null if no webhooks.
        /// </summary>
        public Uri NotificationUrl { get; set; }

        internal INameResolver appSettings;

        /// <summary>
        /// Initialize the O365 binding extension
        /// </summary>
        /// <param name="context">Context containing info relevant to this extension</param>
        public void Initialize(ExtensionConfigContext context)
        {
            var config = context.Config;
            this.appSettings = config.NameResolver;

            // Set up token extension; handles auth (only providers supported by Easy Auth)
            this.tokenExtension = new AuthTokenExtensionConfig();
            this.tokenExtension.InitializeAllExceptRules(context);
            //config.AddExtension(this.tokenExtension);

            // Set up logging
            this.Log = context.Trace;

            // Infer a blank Notification URL from the appsettings.
            if (this.NotificationUrl == null)
            {
                this.NotificationUrl = context.GetWebhookHandler();
            }

            var converter = new Converters(this);

            // Extend token attribute to retrieve [authenticated] GraphServiceClient
            //this.tokenExtension.TokenRule.BindToInput<GraphServiceClient>(converter);

            // Webhooks
            var webhookSubscriptionRule = context.AddBindingRule<GraphWebhookSubscriptionAttribute>();

            webhookSubscriptionRule.BindToInput<Subscription[]>(converter);
            webhookSubscriptionRule.BindToInput<string[]>(converter);
            webhookSubscriptionRule.BindToCollector<string>(converter.CreateCollector);

            string appSettingBYOBTokenMap = appSettings.Resolve(O365Constants.AppSettingBYOBTokenMap);
            this.subscriptionStore = new WebhookSubscriptionStore(appSettingBYOBTokenMap);
            this.webhookTriggerProvider = new WebhookTriggerBindingProvider();
            context.AddBindingRule<GraphWebhookTriggerAttribute>().BindToTrigger(this.webhookTriggerProvider);

            // OneDrive
            var OneDriveRule = context.AddBindingRule<OneDriveAttribute>();

            // OneDrive inputs
            OneDriveRule.BindToInput<byte[]>(converter);
            OneDriveRule.BindToInput<string>(converter);
            OneDriveRule.BindToInput<Stream>(converter);
            OneDriveRule.BindToInput<DriveItem>(converter);

            // OneDrive Outputs
            OneDriveRule.AddConverter<byte[], Stream>(OneDriveClient.CreateStream);
            OneDriveRule.BindToCollector<Stream>(converter.CreateCollector);

            // Excel
            var ExcelRule = context.AddBindingRule<ExcelAttribute>();

            // Excel Outputs
            ExcelRule.AddConverter<object[][], JObject>(ExcelClient.CreateRows);
            ExcelRule.AddConverter<List<OpenType>, JObject>(typeof(GenericConverter<>)); // used to append/update lists of POCOs
            ExcelRule.AddConverter<OpenType, JObject>(typeof(GenericConverter<>)); // used to append/update arrays of POCOs
            ExcelRule.BindToCollector<JObject>(converter.CreateCollector);
            ExcelRule.BindToCollector<JObject>(typeof(POCOConverter<>));

            // Excel Inputs
            ExcelRule.BindToInput<string[][]>(converter);
            ExcelRule.BindToInput<WorkbookTable>(converter);
            ExcelRule.BindToInput<List<OpenType>>(typeof(POCOConverter<>), this);
            ExcelRule.BindToInput<OpenType>(typeof(POCOConverter<>), this);

            // Outlook
            var OutlookRule = context.AddBindingRule<OutlookAttribute>();

            // Outlook Outputs
            OutlookRule.AddConverter<JObject, Message>(OutlookClient.CreateMessage);
            OutlookRule.AddConverter<string, Message>(OutlookClient.CreateMessage);
            OutlookRule.BindToCollector<Message>(converter.CreateCollector);
        }

        /// <summary>
        /// Retrieve audience from raw JWT
        /// </summary>
        /// <param name="rawToken">JWT</param>
        /// <returns>Token audience</returns>
        public static string GetTokenOID(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var oidClaim = jwt.Claims.FirstOrDefault(claim => claim.Type == "oid");
            if (oidClaim == null)
            {
                throw new InvalidOperationException("The graph token is missing an oid. Check your Microsoft Graph binding configuration.");
            }
            return oidClaim.Value;
        }

        /// <summary>
        /// Given a JWT, return the list of scopes in alphabetical order
        /// </summary>
        /// <param name="rawToken">raw JWT</param>
        /// <returns>string of scopes in alphabetical order, separated by a space</returns>
        public static string GetTokenOrderedScopes(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var stringScopes = jwt.Claims.FirstOrDefault(claim => claim.Type == "scp")?.Value;
            if(stringScopes == null)
            {
                throw new InvalidOperationException("The graph token has no scopes. Ensure your application is properly configured to access the Microsoft Graph.");
            }
            var scopes = stringScopes.Split(' ');
            Array.Sort(scopes);
            return string.Join(" ", scopes);
        }

        /// <summary>
        /// Retrieve integer token expiration date
        /// </summary>
        /// <param name="rawToken">raw JWT</param>
        /// <returns>parsed expiration date</returns>
        public static int GetTokenExpirationDate(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var stringTime = jwt.Claims.FirstOrDefault(claim => claim.Type == "exp").Value;
            int result;
            if (int.TryParse(stringTime, out result))
            {
                return result;
            } else
            {
                return -1;
            }
        }

        /// <summary>
        /// Upon receiving webhook trigger data, process it
        /// </summary>
        /// <param name="data">Data from MS Graph -> triggers webhook fx</param>
        /// <returns>Task awaiting result of pushing webhook data</returns>
        internal async Task OnWebhookReceived(WebhookTriggerData data)
        {
            await this.webhookTriggerProvider.PushDataAsync(data);
        }

        /// <summary>
        /// Hydrate GraphServiceClient from a moniker (serialized TokenAttribute)
        /// </summary>
        /// <param name="moniker">string representing serialized TokenAttribute</param>
        /// <returns>Authenticated GraphServiceClient</returns>
        public virtual async Task<GraphServiceClient> GetMSGraphClientFromUserIdAsync(string userId)
        {
            var attr = new TokenAttribute
            {
                UserId = userId,
                Resource = O365Constants.GraphBaseUrl,
                Identity = TokenIdentityMode.UserFromId,
            };

            return await this.GetMSGraphClientAsync(attr);
        }

        /// <summary>
        /// Either retrieve existing GSC or create a new one
        /// GSCs are cached using a combination of the user's principal ID and the scopes of the token used to authenticate
        /// </summary>
        /// <param name="attribute">Token attribute with either principal ID or ID token</param>
        /// <returns>Authenticated GSC</returns>
        public async Task<GraphServiceClient> GetMSGraphClientAsync(TokenAttribute attribute)
        {
            string token = await this.tokenExtension.GetAccessTokenAsync(attribute);
            string principalId = GetTokenOID(token);

            var key = string.Concat(principalId, " ", GetTokenOrderedScopes(token));

            CachedClient client = null;

            // Check to see if there already exists a GSC associated with this principal ID and the token scopes.
            if (this.clients.TryGetValue(key, out client))
            {
                // Check if token is expired
                if (client.expirationDate < DateTimeOffset.Now.ToUnixTimeSeconds())
                {
                    // Need to update the client's token & expiration date
                    // $$ todo -- just reset token instead of whole new authentication provider?
                    client.client.AuthenticationProvider = new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

                            return Task.CompletedTask;
                        });
                    client.expirationDate = GetTokenExpirationDate(token);
                }

                return client.client;
            }
            else
            {
                client = new CachedClient
                {
                    client = new GraphServiceClient(
                        new DelegateAuthenticationProvider(
                            (requestMessage) =>
                            {
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                                return Task.CompletedTask;
                            })),
                    expirationDate = GetTokenExpirationDate(token),
                };
                this.clients.TryAdd(key, client);
                return client.client;
            }
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
            var handler = new GraphWebhookSubscriptionHandler(this);
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
                    JObject jsonContent = new JObject();

                    // T[] -> JArray
                    JArray rowData = JArray.FromObject(input);

                    jsonContent[O365Constants.ValuesKey] = rowData;

                    // Set rows, columns needed if updating entire worksheet
                    jsonContent[O365Constants.RowsKey] = rowData.Count();

                    // No exception -- array is rectangular by default
                    jsonContent[O365Constants.ColsKey] = rowData.First.Count();

                    // Set POCO key to indicate that the values need to be ordered to match the header of the existing table
                    jsonContent[O365Constants.POCOKey] = true;

                    return jsonContent;
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
        public class POCOConverter<T> : IAsyncConverter<ExcelAttribute, T[]>, IAsyncConverter<ExcelAttribute, List<T>>
            where T : new()
        {
            private readonly MicrosoftGraphExtensionConfig parent;

            /// <summary>
            /// Initializes a new instance of the <see cref="POCOConverter{T}"/> class.
            /// </summary>
            /// <param name="parent">O365Extension to which the result of the request for data will be returned</param>
            public POCOConverter(MicrosoftGraphExtensionConfig parent)
            {
                this.parent = parent;
            }

            async Task<List<T>> IAsyncConverter<ExcelAttribute, List<T>>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var client = await this.parent.GetMSGraphClientAsync(input);
                var result = await client.GetExcelRangePOCOListAsync<T>(input);
                return result;
            }

            async Task<T[]> IAsyncConverter<ExcelAttribute, T[]>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var client = await this.parent.GetMSGraphClientAsync(input);
                var result = client.GetExcelRangePOCOAsync<T>(input);
                return await result;
            }

            public IAsyncCollector<JObject> CreateCollector(ExcelAttribute attr)
            {
                GraphServiceClient client = this.parent.GetMSGraphClientAsync(attr).Result;
                return new ExcelAsyncCollector(client, attr);
            }
        }

        /// <summary>
        /// Used for input bindings; Attribute -> Input type
        /// </summary>
        public class Converters :
            IAsyncConverter<ExcelAttribute, string[][]>,
            IAsyncConverter<ExcelAttribute, WorkbookTable>,
            IAsyncConverter<OneDriveAttribute, byte[]>,
            IAsyncConverter<OneDriveAttribute, string>,
            IAsyncConverter<OneDriveAttribute, Stream>,
            IAsyncConverter<OneDriveAttribute, DriveItem>,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, Subscription[]>,
            IAsyncConverter<GraphWebhookSubscriptionAttribute, string[]>
        {
            private readonly MicrosoftGraphExtensionConfig _parent;

            public Converters(MicrosoftGraphExtensionConfig parent)
            {
                _parent = parent;
            }

            public IAsyncCollector<JObject> CreateCollector(ExcelAttribute attr)
            {
                GraphServiceClient client = _parent.GetMSGraphClientAsync(attr).Result;
                return new ExcelAsyncCollector(client, attr);
            }

            public IAsyncCollector<Stream> CreateCollector(OneDriveAttribute attr)
            {
                GraphServiceClient client = _parent.GetMSGraphClientAsync(attr).Result;
                return new OneDriveAsyncCollector(client, attr);
            }

            public IAsyncCollector<Message> CreateCollector(OutlookAttribute attr)
            {
                GraphServiceClient client = _parent.GetMSGraphClientAsync(attr).Result;
                return new OutlookAsyncCollector(client, attr);
            }

            public IAsyncCollector<string> CreateCollector(GraphWebhookSubscriptionAttribute attr)
            {
                return new GraphWebhookSubscriptionAsyncCollector(_parent, attr);
            }

            async Task<string[][]> IAsyncConverter<ExcelAttribute, string[][]>.ConvertAsync(ExcelAttribute attr, CancellationToken cancellationToken)
            {
                GraphServiceClient client = _parent.GetMSGraphClientAsync(attr).Result;
                var result = await client.GetExcelRangeAsync(attr);
                return result;
            }

            async Task<WorkbookTable> IAsyncConverter<ExcelAttribute, WorkbookTable>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var client = await _parent.GetMSGraphClientAsync(input);

                var result = await client.GetExcelTable(input);
                return result;
            }

            async Task<byte[]> IAsyncConverter<OneDriveAttribute, byte[]>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var client = await _parent.GetMSGraphClientAsync(input);

                var result = await client.GetOneDriveContentsAsync(input);

                return result;
            }

            async Task<string> IAsyncConverter<OneDriveAttribute, string>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var graphClient = await _parent.GetMSGraphClientAsync(input);

                var byteArray = await graphClient.GetOneDriveContentsAsync(input);

                return Encoding.UTF8.GetString(byteArray);
            }

            async Task<Stream> IAsyncConverter<OneDriveAttribute, Stream>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var client = await _parent.GetMSGraphClientAsync(input);

                return await client.GetOneDriveContentStreamAsync(input);
            }

            async Task<DriveItem> IAsyncConverter<OneDriveAttribute, DriveItem>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
            {
                var client = await _parent.GetMSGraphClientAsync(input);
                return await client.GetOneDriveContentDriveItemAsync(input);
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
                IEnumerable<WebhookSubscriptionStore.SubscriptionEntry> subscriptionEntries = await _parent.subscriptionStore.GetAllSubscriptionsAsync();
                if (TokenIdentityMode.UserFromRequest.ToString().Equals(attribute.Filter, StringComparison.OrdinalIgnoreCase))
                {
                    var dummyTokenAttribute = new TokenAttribute()
                    {
                        Resource = O365Constants.GraphBaseUrl,
                        Identity = TokenIdentityMode.UserFromToken,
                        UserToken = attribute.UserToken,
                        IdentityProvider = "AAD",
                    };
                    var graph = await _parent.GetMSGraphClientAsync(dummyTokenAttribute);
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
