// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2, PublicKey=0024000004800000940000000602000000240000525341310004000001000100cd1dabd5a893b40e75dc901fe7293db4a3caf9cd4d3e3ed6178d49cd476969abe74a9e0b7f4a0bb15edca48758155d35a4f05e6e852fff1b319d103b39ba04acbadd278c2753627c95e1f6f6582425374b92f51cca3deb0d2aab9de3ecda7753900a31f70a236f163006beefffe282888f85e3c76d1205ec7dfef7fa472a17b1")]
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
