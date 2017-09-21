// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Executors;
    using Microsoft.Azure.WebJobs.Host.Listeners;
    using Microsoft.Azure.WebJobs.Host.Protocols;
    using Microsoft.Azure.WebJobs.Host.Triggers;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    internal class WebhookTriggerBindingProvider :
         ITriggerBindingProvider
    {
        private Dictionary<string, WebhookTriggerListener> _listenerPerDataType = new Dictionary<string, WebhookTriggerListener>();
        private WebhookTriggerListener _generalListener;

        public WebhookTriggerBindingProvider()
        {
            _generalListener = null;
        }

        internal void AddListener(GraphWebhookTriggerAttribute attribute, WebhookTriggerListener listener)
        {
            if (attribute.ResourceType == null)
            {
                if (_generalListener == null) {
                    throw new InvalidOperationException($"Cannot have multiple graph webhook listeners without types.");
                }

                _generalListener = listener;
            }
            else
            {
                if (_listenerPerDataType.ContainsKey(attribute.ResourceType))
                {
                    throw new InvalidOperationException($"Cannot have more than one graph webhook listener for {attribute.ResourceType}");
                }
                _listenerPerDataType.Add(attribute.ResourceType, listener);
            }
        }

        internal void StopListeners()
        {
            _listenerPerDataType.Clear();
            _generalListener = null;
        }

        // Called by external webhook to push data. 
        public async Task PushDataAsync(WebhookTriggerData data)
        {
            WebhookTriggerListener dataListener = _generalListener;

            if (_listenerPerDataType.ContainsKey(data.odataType))
            {
                dataListener = _listenerPerDataType[data.odataType];
            }

            if (dataListener != null)
            {
                var exec = dataListener.Executor;

                TriggeredFunctionData input = new TriggeredFunctionData
                {
                    ParentId = null,
                    TriggerValue = data,
                };
                FunctionResult result = await exec.TryExecuteAsync(input, CancellationToken.None);
            }
        }

        public async Task<ITriggerBinding> TryCreateAsync(TriggerBindingProviderContext context)
        {
            ParameterInfo parameter = context.Parameter;

            GraphWebhookTriggerAttribute attribute = parameter.GetCustomAttribute<GraphWebhookTriggerAttribute>(inherit: false);

            ITriggerBinding binding = null;

            if (attribute != null)
            {
                binding = new MyTriggerBinding(this, attribute, parameter);
            }

            return binding;
        }

        // Per parameter. 
        class MyTriggerBinding : ITriggerBinding
        {
            public const string AuthName = "auth";

            private readonly GraphWebhookTriggerAttribute _attribute;
            private readonly WebhookTriggerBindingProvider _parent;
            private readonly ParameterInfo _parameter;

            public MyTriggerBinding(WebhookTriggerBindingProvider parent, GraphWebhookTriggerAttribute attribute, ParameterInfo parameter)
            {
                _attribute = attribute;
                _parent = parent;
                _parameter = parameter;
            }

            public IReadOnlyDictionary<string, Type> BindingDataContract
            {
                get
                {
                    return new Dictionary<string, Type>()
                    {
                        { AuthName,  typeof(string) }, // $$$ share
                        { "MSGraphWebhook", typeof(string) },
                    };
                }
            }

            public Type TriggerValueType
            {
                get
                {
                    return typeof(WebhookTriggerData);
                }
            }

            private static async Task<string> GetTokenFromGraphClientAsync(GraphServiceClient client)
            {
                if (client == null)
                {
                    return null;
                }

                HttpRequestMessage request = new HttpRequestMessage
                {
                    Method = HttpMethod.Get,
                    RequestUri = new Uri(client.BaseUrl),
                };

                await client.AuthenticationProvider.AuthenticateRequestAsync(request);
                return request.Headers.Authorization.Parameter;
            }

            public Task<ITriggerData> BindAsync(object value, ValueBindingContext context)
            {
                WebhookTriggerData data = (WebhookTriggerData)value;
                var bindingData = new Dictionary<string, object>();

                bindingData[AuthName] = GetTokenFromGraphClientAsync(data.client);

                JObject raw = data.Payload;
                var userObject = raw.ToObject(_parameter.ParameterType);

                IValueProvider valueProvider = new ObjectValueProvider(userObject);
                var triggerData = new TriggerData(valueProvider, bindingData);
                return Task.FromResult<ITriggerData>(triggerData);
            }

            public Task<IListener> CreateListenerAsync(ListenerFactoryContext context)
            {
                var listener = new WebhookTriggerListener(context, _parent, _attribute);
                return Task.FromResult<IListener>(listener);
            }

            public ParameterDescriptor ToParameterDescriptor()
            {
                return new ParameterDescriptor
                {
                    Name = _parameter.Name,
                };
            }
        }

        class ObjectValueProvider : IValueProvider
        {
            private object value;

            public ObjectValueProvider(object value)
            {
                this.value = value;
            }

            public Type Type
            {
                get
                {
                    return this.value.GetType();
                }
            }

            public Task<object> GetValueAsync()
            {
                return Task.FromResult<object>(this.value);
            }

            public string ToInvokeString()
            {
                return this.value.ToString();
            }
        }
    } // WebhookTriggerBindingProvider
}
