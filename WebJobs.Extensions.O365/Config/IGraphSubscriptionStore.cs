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
        /// <summary>
        /// Saves the subscription in the token store, along with the user id for the subscription
        /// </summary>
        /// <param name="subscription">The subscription being saved</param>
        /// <param name="userId">The user id of the user who the subscription belongs to</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        Task SaveSubscriptionEntryAsync(Subscription subscription, string userId);

        /// <summary>
        /// Retrieves all subscription entries in the store.
        /// </summary>
        /// <returns>A task that respresents the asynchronous operation. The task's result contains the 
        /// subscription entries.</returns>
        Task<SubscriptionEntry[]> GetAllSubscriptionsAsync();

        /// <summary>
        /// Retrieves the subscription entry with the given subscription id.
        /// Returns null if the subscriptionId does not match any subscriptions in the store
        /// </summary>
        /// <param name="subscriptionId">The subscription id of the entry to retrieve</param>
        /// <returns>A task that respresents the asynchronous operation. The task's result contains the 
        /// subscription entry with the subscription Id provided</returns>
        Task<SubscriptionEntry> GetSubscriptionEntryAsync(string subscriptionId);

        /// <summary>
        /// Deletes the subscription entry in the store with the given subscription Id.
        /// NoOps if the subscriptionId does not match any subscriptions in the store
        /// </summary>
        /// <param name="subscriptionId">The subscription id of the entry to delete</param>
        /// <returns>A task that represents the asynchronous operation</returns>
        Task DeleteAsync(string subscriptionId);
    }
}
