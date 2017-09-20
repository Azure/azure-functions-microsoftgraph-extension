// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    /// <summary>
    /// Helper class for calling onto (MS) OneDrive Graph
    /// </summary>
    internal static class OneDriveClient
    {
        /// <summary>
        /// Retrieve contents of OneDrive file
        /// </summary>
        /// <param name="client"></param>
        /// <param name="graphClient"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public static async Task<byte[]> GetOneDriveContentsAsync(this GraphServiceClient client, OneDriveAttribute attr)
        {
            // How to download from OneDrive: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_downloadcontent
            // GET https://graph.microsoft.com/v1.0/me/drive/root:/test1/hi.txt:/content HTTP/1.1
            bool isShare = attr.Path.StartsWith("https://");
            if (isShare)
            {
                // TODO: Move this to GraphServiceClient

                // Download via a Share URL
                var shareToken = UrlToSharingToken(attr.Path);
                var response = await client.Shares[shareToken].Root.Content.Request().GetAsync();

                using (MemoryStream ms = new MemoryStream())
                {
                    await response.CopyToAsync(ms);
                    return ms.ToArray();
                }
            }
            else
            {
                // Retrieve stream of OneDrive item
                var stream = await client
                    .Me
                    .Drive
                    .Root
                    .ItemWithPath(attr.Path)
                    .Content
                    .Request()
                    .GetAsync();

                // Convert to Memory Stream
                MemoryStream ms = new MemoryStream();
                await stream.CopyToAsync(ms);

                // Convert to Byte Array
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Retrieve contents of OneDrive file
        /// </summary>
        /// <param name="client">Authenticated Graph Service Client used to retrieve file</param>
        /// <param name="attr">Attribute with necessary data (e.g. path)</param>
        /// <returns>Stream of file content</returns>
        public static async Task<Stream> GetOneDriveContentStreamAsync(this GraphServiceClient client, OneDriveAttribute attr)
        {
            // Retrieve stream of OneDrive item
            return await client
                .Me
                .Drive
                .Root
                .ItemWithPath(attr.Path)
                .Content
                .Request()
                .GetAsync();
        }

        public static async Task<DriveItem> GetOneDriveContentDriveItemAsync(this GraphServiceClient client, OneDriveAttribute attr)
        {
            // Retrieve OneDrive item
            return await client
                .Me
                .Drive
                .Root
                .ItemWithPath(attr.Path)
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
        public static async Task<DriveItem> UploadOneDriveItemAsync(this GraphServiceClient graphClient, OneDriveAttribute attr, Stream fileStream)
        {
            return await graphClient
                .Me
                .Drive
                .Root
                .ItemWithPath(attr.Path)
                .Content
                .Request()
                .PutAsync<DriveItem>(fileStream);
        }

        /// <summary>
        /// Uploads new OneDrive Item OR updates existing OneDrive Item
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="attr"></param>
        /// <param name="byteArray">Byte array to upload</param>
        /// <returns>DriveItem representing newly added/updated item</returns>
        public static async Task<DriveItem> UploadOneDriveItemAsync(this GraphServiceClient graphClient, OneDriveAttribute attr, byte[] byteArray)
        {

            return await graphClient
                .Me
                .Drive
                .Root
                .ItemWithPath(attr.Path)
                .Content
                .Request()
                .PutAsync<DriveItem>(CreateStream(byteArray));
        }

        public static Stream CreateStream(byte[] byteArray)
        {
            return new MemoryStream(byteArray);
        }

        // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/shares_get#transform-a-sharing-url
        public static string UrlToSharingToken(string inputUrl)
        {
            var base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(inputUrl));
            return "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
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