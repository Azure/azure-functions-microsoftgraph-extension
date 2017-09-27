// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System.IO;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    internal class OneDriveService
    {
        private IOneDriveClient _client;

        public OneDriveService(IOneDriveClient client)
        {
            _client = client;
        }

        /// <summary>
        /// Retrieve contents of OneDrive file
        /// </summary>
        /// <param name="client"></param>
        /// <param name="graphClient"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public async Task<byte[]> GetOneDriveContentsAsByteArrayAsync(OneDriveAttribute attr)
        {
            // How to download from OneDrive: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_downloadcontent
            // GET https://graph.microsoft.com/v1.0/me/drive/root:/test1/hi.txt:/content HTTP/1.1
            bool isShare = attr.Path.StartsWith("https://");
            if (isShare)
            {
                // TODO: Move this to GraphServiceClient

                // Download via a Share URL
                var shareToken = UrlToSharingToken(attr.Path);
                var response = await _client.GetOneDriveContentStreamFromShareAsync(shareToken);

                using (MemoryStream ms = new MemoryStream())
                {
                    await response.CopyToAsync(ms);
                    return ms.ToArray();
                }
            }
            else
            {
                // Retrieve stream of OneDrive item
                var stream = await _client.GetOneDriveContentStreamAsync(attr.Path);

                // Convert to Memory Stream
                MemoryStream ms = new MemoryStream();
                await stream.CopyToAsync(ms);

                // Convert to Byte Array
                return ms.ToArray();
            }
        }

        public async Task<Stream> GetOneDriveContentsAsStreamAsync(OneDriveAttribute attr)
        {
            return await _client.GetOneDriveContentStreamAsync(attr.Path);
        }

        public async Task<DriveItem> GetOneDriveItemAsync(OneDriveAttribute attr)
        {
            return await _client.GetOneDriveItemAsync(attr.Path);
        }

        public async Task<DriveItem> UploadOneDriveContentsAsync(OneDriveAttribute attr, Stream fileStream)
        {
            return await _client.UploadOneDriveItemAsync(attr.Path, fileStream);
        }

        // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/shares_get#transform-a-sharing-url
        private static string UrlToSharingToken(string inputUrl)
        {
            var base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(inputUrl));
            return "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
        }


        public static Stream CreateStream(byte[] byteArray)
        {
            return new MemoryStream(byteArray);
        }
    }
}
