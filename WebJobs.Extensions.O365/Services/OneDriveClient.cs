// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    /// <summary>
    /// Helper class for calling onto (MS) OneDrive Graph
    /// </summary>
    internal class OneDriveClient : IOneDriveClient
    {
        private Task<IGraphServiceClient> _client;

        public OneDriveClient(Task<IGraphServiceClient> client)
        {
            _client = client;
        }

        /// <summary>
        /// Retrieve contents of OneDrive file
        /// </summary>
        /// <param name="client">Authenticated Graph Service Client used to retrieve file</param>
        /// <param name="attr">Attribute with necessary data (e.g. path)</param>
        /// <returns>Stream of file content</returns>
        public async Task<Stream> GetOneDriveContentStreamAsync(string path)
        {
            // Retrieve stream of OneDrive item
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Content
                .Request()
                .GetAsync();
        }

        public async Task<DriveItem> GetOneDriveItemAsync(string path)
        {
            // Retrieve OneDrive item
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Request()
                .GetAsync();
        }

        public async Task<Stream> GetOneDriveContentStreamFromShareAsync(string shareToken)
        {
            return await (await _client)
                .Shares[shareToken]
                .Root
                .Content
                .Request()
                .GetAsync();
        }

        /// <summary>
        /// Uploads new OneDrive Item OR updates existing OneDrive Item
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="attr"></param>
        /// <param name="fileStream">Stream of input to be uploaded</param>
        /// <returns>Drive item representing newly added/updated item</returns>
        public async Task<DriveItem> UploadOneDriveItemAsync(string path, Stream fileStream)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Content
                .Request()
                .PutAsync<DriveItem>(fileStream);
        }

        class GetRootModel
        {
            [JsonProperty("@microsoft.graph.downloadUrl")]
            public string DownloadUrl { get; set; }

            public string name { get; set; }

            public int size { get; set; }
        }
    }
}