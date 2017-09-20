// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Host.Executors;
    using Microsoft.Azure.WebJobs.Host.Listeners;

    // Per-function.
    // The listener is "passive". It just tracks the context invoker used for calling functions.
    // This receives context that can invoke the function
    internal class WebhookTriggerListener : IListener
    {
        // The context contains an invoker that can call the user function
        private readonly ListenerFactoryContext _context;
        private readonly WebhookTriggerBindingProvider _parent;
        private readonly GraphWebhookTriggerAttribute _attribute;

        public WebhookTriggerListener(ListenerFactoryContext context, WebhookTriggerBindingProvider parent, GraphWebhookTriggerAttribute attribute)
        {
            this._context = context;
            this._parent = parent;
            this._attribute = attribute;
        }

        public void Cancel()
        {
        }

        public void Dispose()
        {
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            this._parent.AddListener(this._attribute, this);
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            this._parent.StopListeners();

            return Task.CompletedTask;
        }

        public ITriggeredFunctionExecutor Executor => this._context.Executor;
    }
}