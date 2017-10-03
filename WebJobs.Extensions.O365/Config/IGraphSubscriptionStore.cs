// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]
namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    using System.Threading.Tasks;
    using Microsoft.Graph;

    internal interface IGraphSubscriptionStore
    {
        Task SaveSubscriptionEntryAsync(Subscription subscription, string userId);

        Task<SubscriptionEntry[]> GetAllSubscriptionsAsync();

        Task<SubscriptionEntry> GetSubscriptionEntryAsync(string subId);

        Task DeleteAsync(string subscriptionId);
    }
}
