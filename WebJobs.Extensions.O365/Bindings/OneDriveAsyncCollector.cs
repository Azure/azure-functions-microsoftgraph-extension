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
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    /// <summary>
    /// Collector class used to accumulate and then dispatch requests to MS Graph related to OneDrive
    /// </summary>
    internal class OneDriveAsyncCollector : IAsyncCollector<Stream>
    {
        private readonly OneDriveService _service;
        private readonly OneDriveAttribute _attribute;
        private readonly Collection<Stream> _fileStreams;

        /// <summary>
        /// Initializes a new instance of the <see cref="OneDriveAsyncCollector"/> class.
        /// </summary>
        /// <param name="client">GraphServiceClient used to make calls to MS Graph</param>
        /// <param name="attribute">OneDriveAttribute containing necessary info about file</param>
        public OneDriveAsyncCollector(OneDriveService service, OneDriveAttribute attribute)
        {
            _service = service;
            _attribute = attribute;
            _fileStreams = new Collection<Stream>();
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

            _fileStreams.Add(item);
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
            foreach (var stream in _fileStreams)
            {
                await _service.UploadOneDriveContentsAsync(_attribute, stream);
            }

            this._fileStreams.Clear();
        }
    }
}