// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Graph;
    using Newtonsoft.Json;
    using File = System.IO.File;

    /// <summary>
    /// Store mapping from webhook subscription IDs to a token.
    /// </summary>
    internal class WebhookSubscriptionStore : IGraphSubscriptionStore
    {
        private string root; // @"C:\temp\sub";

        private FileLock _fileLock;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebhookSubscriptionStore"/> class.
        /// Find webhook token cache path
        /// If it doesn't exist, create directory
        /// </summary>
        /// <param name="home">Value of app setting used to det. webhook token cache path</param>
        public WebhookSubscriptionStore(string home)
        {
            this.root = home ?? O365Constants.DefaultBYOBLocation;
            _fileLock = new FileLock();
            _fileLock.PerformWriteIO(this.root, () => CreateRootDirectory(this.root));
        }

        private static void CreateRootDirectory(string root)
        {
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
        }

        private string GetSubscriptionPath(string subscriptionId)
        {
            return Path.Combine(this.root, subscriptionId);
        }

        private string GetSubscriptionPath(Subscription subscription)
        {
            return GetSubscriptionPath(subscription.Id);
        }

        public async Task SaveSubscriptionEntryAsync(Subscription subscription, string userId)
        {
            var entry = new SubscriptionEntry
            {
                Subscription = subscription,
                UserId = userId,
            };
            var subPath = this.GetSubscriptionPath(subscription);
            var jsonString = JsonConvert.SerializeObject(entry);
            await _fileLock.PerformWriteIOAsync(subPath, () => File.WriteAllText(subPath, jsonString));
        }

        public async Task<SubscriptionEntry[]> GetAllSubscriptionsAsync()
        {
            var subscriptionPaths = await _fileLock.PerformReadIOAsync<IEnumerable<string>>(this.root, Directory.EnumerateFiles);
            var entryTasks = subscriptionPaths.Select(path => this.GetAsyncFromPath(path));
            return await Task.WhenAll(entryTasks);
        }

        private async Task<SubscriptionEntry> GetAsyncFromPath(string path)
        {
            var contents = await _fileLock.PerformReadIOAsync<string>(path, File.ReadAllText);
            var entry = JsonConvert.DeserializeObject<SubscriptionEntry>(contents);
            return entry;
        }

        public async Task<SubscriptionEntry> GetSubscriptionEntryAsync(string subId)
        {
            string path = this.GetSubscriptionPath(subId);
            return await this.GetAsyncFromPath(path);
        }

        /// <summary>
        /// Delete a single subscription entry
        /// </summary>
        /// <param name="entry">Subscription entry to be deleted</param>
        public async Task DeleteAsync(string subscriptionId)
        {
            var path = this.GetSubscriptionPath(subscriptionId);
            await _fileLock.PerformWriteIOAsync(path, () => DeleteFileIfExists(path));
        }

        private void DeleteFileIfExists(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private class FileLock
        {
            private readonly ConcurrentDictionary<string, object> _locks = new ConcurrentDictionary<string, object>();

            public void PerformWriteIO(string path, Action ioAction)
            {
                lock (_locks.GetOrAdd(path, new object()))
                {
                    try
                    {
                        ioAction();
                    }
                    finally
                    {
                        object removedLock;
                        _locks.TryRemove(path, out removedLock);
                    }
                }
            }

            public async Task PerformWriteIOAsync(string path, Action ioAction)
            {
                PerformWriteIO(path, ioAction);
            }

            public async Task<T> PerformReadIOAsync<T>(string path, Func<string, T> ioAction)
            {
                lock (_locks.GetOrAdd(path, new object()))
                {
                    T result = default(T);
                    try
                    {
                        result = ioAction(path);
                    }
                    finally
                    {
                        object removedLock;
                        _locks.TryRemove(path, out removedLock);
                    }

                    return result;
                }
            }
        }
    }
}
