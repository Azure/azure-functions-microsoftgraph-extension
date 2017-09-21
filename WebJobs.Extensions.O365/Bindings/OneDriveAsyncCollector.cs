/// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;

    /// <summary>
    /// Collector class used to accumulate and then dispatch requests to MS Graph related to OneDrive
    /// </summary>
    internal class OneDriveAsyncCollector : IAsyncCollector<Stream>
    {
        private readonly GraphServiceClient client;
        private readonly OneDriveAttribute attribute;
        private readonly Collection<Stream> fileStreams = new Collection<Stream>();

        /// <summary>
        /// Initializes a new instance of the <see cref="OneDriveAsyncCollector"/> class.
        /// </summary>
        /// <param name="client">GraphServiceClient used to make calls to MS Graph</param>
        /// <param name="attribute">OneDriveAttribute containing necessary info about file</param>
        public OneDriveAsyncCollector(GraphServiceClient client, OneDriveAttribute attribute)
        {
            this.client = client;
            this.attribute = attribute;
        }

        /// <summary>
        /// Add a stream representing a file to the list of objects needing to be processed
        /// </summary>
        /// <param name="item">Stream representing OneDrive file</param>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>Task whose resolution results in Stream being added to the collector</returns>
        public Task AddAsync(Stream item, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (item == null)
            {
                throw new ArgumentNullException("No stream found");
            }

            this.fileStreams.Add(item);
            return Task.CompletedTask;
        }

        /// <summary>
        /// Execute methods associated with file streams
        /// </summary>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>Task representing the file streams being uploaded</returns>
        public async Task FlushAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            // Upload all files
            foreach (var stream in this.fileStreams)
            {
                await this.client.UploadOneDriveItemAsync(this.attribute, stream);
            }

            this.fileStreams.Clear();
        }
    }
}