// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    internal class OneDriveWriteStream : Stream
    {
        private readonly IGraphServiceClient _client;
        private readonly Stream _stream;
        private readonly string _path;

        public OneDriveWriteStream(IGraphServiceClient client, string path)
        {
            _client = client;
            _path = path;
            _stream = new MemoryStream();
        }

        public override bool CanRead => false;

        public override bool CanSeek =>  false;

        public override bool CanWrite => true;

        public override long Length => _stream.Length;

        public override long Position { get => _stream.Position; set => throw new NotSupportedException("This stream cannot seek"); }

        public override void Flush()
        {
            //NO-OP since flush must be idempotent, and the upload operation is not
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException("This stream does not support read operations.");
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException("This stream does not support seek operations");
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException("This stream does not support seek operations");
        }

        public override void Close()
        {
            Task.Run(() => _client.UploadOneDriveItemAsync(_path, _stream)).GetAwaiter().GetResult();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _stream.Write(buffer, offset, count);
        }
    }
}
