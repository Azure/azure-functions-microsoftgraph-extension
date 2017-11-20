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

    class OneDriveConverter :
        IAsyncConverter<OneDriveAttribute, byte[]>,
        IAsyncConverter<OneDriveAttribute, string>,
        IAsyncConverter<OneDriveAttribute, Stream>,
        IAsyncConverter<OneDriveAttribute, DriveItem>
    {
        private readonly ServiceManager _serviceManager;

        public OneDriveConverter(ServiceManager serviceManager)
        {
            _serviceManager = serviceManager;
        }

        async Task<byte[]> IAsyncConverter<OneDriveAttribute, byte[]>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            var service = await _serviceManager.GetOneDriveService(input);
            return await service.GetOneDriveContentsAsByteArrayAsync(input);
        }

        async Task<string> IAsyncConverter<OneDriveAttribute, string>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            var service = await _serviceManager.GetOneDriveService(input);
            var byteArray = await service.GetOneDriveContentsAsByteArrayAsync(input);
            return Encoding.UTF8.GetString(byteArray);
        }

        async Task<Stream> IAsyncConverter<OneDriveAttribute, Stream>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            var service = await _serviceManager.GetOneDriveService(input);
            return await service.GetOneDriveContentsAsStreamAsync(input);
        }

        async Task<DriveItem> IAsyncConverter<OneDriveAttribute, DriveItem>.ConvertAsync(OneDriveAttribute input, CancellationToken cancellationToken)
        {
            var service = await _serviceManager.GetOneDriveService(input);
            return await service.GetOneDriveItemAsync(input);
        }
    }
}
