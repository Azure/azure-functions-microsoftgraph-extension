// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Newtonsoft.Json;
    using File = System.IO.File;
    using System.Collections.Concurrent;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Store mapping from webhook subscription IDs to a token.
    /// </summary>
    internal class WebhookSubscriptionStore
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

        internal async Task SaveSubscriptionEntryAsync(Subscription subscription, string userId)
        {
            var entry = new SubscriptionEntry
            {
                Subscription = subscription,
                UserId = userId,
            };
            var subPath = this.GetSubscriptionPath(subscription);
            var jsonString = JsonConvert.SerializeObject(entry);
            _fileLock.PerformWriteIO(subPath, () => File.WriteAllText(subPath, jsonString));
        }

        /// <summary>
        /// Subscription Entry saved in local storage of app
        /// Location determined by DefaultBYOBLocation or AppSettingBYOBTokenMap
        /// </summary>
        public class SubscriptionEntry
        {
            /// <summary>
            /// Gets or sets subscription ID returned by MS Graph after creation
            /// </summary>
            public Subscription Subscription { get; set; }

            /// <summary>
            /// Gets or sets the user id for the subscription
            /// </summary>
            public string UserId { get; set; } // $$$ Gets an auth token and client
        }

        public async Task<SubscriptionEntry[]> GetAllSubscriptionsAsync()
        {
            var subscriptionPaths = _fileLock.PerformReadIO<IEnumerable<string>>(this.root, Directory.EnumerateFiles);
            var entryTasks = subscriptionPaths.Select(path => this.GetAsyncFromPath(path));
            return await Task.WhenAll(entryTasks);
        }

        private async Task<SubscriptionEntry> GetAsyncFromPath(string path)
        {
            var contents = _fileLock.PerformReadIO<string>(path, File.ReadAllText);
            var entry = JsonConvert.DeserializeObject<SubscriptionEntry>(contents);
            return entry;
        }

        internal async Task<SubscriptionEntry> GetSubscriptionEntryAsync(string subId)
        {
            string path = this.GetSubscriptionPath(subId);
            return await this.GetAsyncFromPath(path);
        }

        /// <summary>
        /// Delete a single subscription entry
        /// </summary>
        /// <param name="entry">Subscription entry to be deleted</param>
        internal async Task DeleteAsync(string subscriptionId)
        {
            var path = this.GetSubscriptionPath(subscriptionId);
            _fileLock.PerformWriteIO(path, () => DeleteFileIfExists(path));
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

            public T PerformReadIO<T>(string path, Func<string, T> ioAction)
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
