// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
using Microsoft.Graph;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    internal class OneDriveService
    {
        private GraphServiceClientManager _clientProvider;

        public OneDriveService(GraphServiceClientManager clientProvider)
        {
            _clientProvider = clientProvider;
        }

        /// <summary>
        /// Retrieve contents of OneDrive file
        /// </summary>
        /// <param name="client"></param>
        /// <param name="graphClient"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public async Task<byte[]> GetOneDriveContentsAsByteArrayAsync(OneDriveAttribute attr, CancellationToken token)
        {
            var response = await GetOneDriveContentsAsStreamAsync(attr, token);

            using (MemoryStream ms = new MemoryStream())
            {
                await response.CopyToAsync(ms);
                return ms.ToArray();
            }
        }

        public Stream ConvertStream(Stream stream, OneDriveAttribute attribute, IGraphServiceClient client)
        {
            if (attribute.Access != FileAccess.Write)
            {
                return stream;
            }
            return new OneDriveWriteStream(client, attribute.Path);
        }

        public async Task<Stream> GetOneDriveContentsAsStreamAsync(OneDriveAttribute attr, CancellationToken token)
        {
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            Stream response;
            // How to download from OneDrive: https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_downloadcontent
            // GET https://graph.microsoft.com/v1.0/me/drive/root:/test1/hi.txt:/content HTTP/1.1
            bool isShare = attr.Path.StartsWith("https://");
            if (isShare)
            {
                // TODO: Move this to GraphServiceClient

                // Download via a Share URL
                var shareToken = UrlToSharingToken(attr.Path);
                response = await client.GetOneDriveContentStreamFromShareAsync(shareToken, token);
            }
            else
            {
                try
                {
                    // Retrieve stream of OneDrive item
                    response = await client.GetOneDriveContentStreamAsync(attr.Path, token);
                } catch
                {
                    //File does not exist, so create new memory stream
                    response = new MemoryStream();
                }


            }

            return ConvertStream(response, attr, client);
        }

        public async Task<DriveItem> GetOneDriveItemAsync(OneDriveAttribute attr, CancellationToken token)
        {
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            return await client.GetOneDriveItemAsync(attr.Path, token);
        }

        public async Task<DriveItem> UploadOneDriveContentsAsync(OneDriveAttribute attr, Stream fileStream, CancellationToken token)
        {
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            return await client.UploadOneDriveItemAsync(attr.Path, fileStream, token);
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
