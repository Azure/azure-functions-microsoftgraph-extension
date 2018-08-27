// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters
{
    using System.IO;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    internal class OneDriveConverter :
        IAsyncConverter<OneDriveAttribute, byte[]>,
        IAsyncConverter<OneDriveAttribute, string>,
        IAsyncConverter<OneDriveAttribute, Stream>,
        IAsyncConverter<OneDriveAttribute, DriveItem>,
        IAsyncConverter<OneDriveAttribute, IAsyncCollector<byte[]>>
    {
        private readonly OneDriveService _service;

        public OneDriveConverter(OneDriveService service)
        {
            _service = service;
        }

        async Task<byte[]> IAsyncConverter<OneDriveAttribute, byte[]>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            return await _service.GetOneDriveContentsAsByteArrayAsync(input, cancellationToken);
        }

        async Task<string> IAsyncConverter<OneDriveAttribute, string>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            var byteArray = await _service.GetOneDriveContentsAsByteArrayAsync(input, cancellationToken);
            return Encoding.UTF8.GetString(byteArray);
        }

        async Task<Stream> IAsyncConverter<OneDriveAttribute, Stream>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            return await _service.GetOneDriveContentsAsStreamAsync(input, cancellationToken);
        }

        async Task<DriveItem> IAsyncConverter<OneDriveAttribute, DriveItem>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            return await _service.GetOneDriveItemAsync(input, cancellationToken);
        }

        public async Task<IAsyncCollector<byte[]>> ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            return new OneDriveAsyncCollector(_service, input);
        }
    }
}
