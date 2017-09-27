// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System.IO;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    internal interface IOneDriveClient
    {
        Task<Stream> GetOneDriveContentStreamAsync(string path);

        Task<DriveItem> GetOneDriveItemAsync(string path);

        Task<Stream> GetOneDriveContentStreamFromShareAsync(string shareToken);

        Task<DriveItem> UploadOneDriveItemAsync(string path, Stream fileStream);
    }
}
