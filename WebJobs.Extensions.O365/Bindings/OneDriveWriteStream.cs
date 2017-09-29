// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.IO;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    internal class OneDriveWriteStream : Stream
    {
        private IGraphServiceClient _client;
        private Stream _stream;
        private string _path;

        public OneDriveWriteStream(IGraphServiceClient client, Stream existingStream, string path)
        {
            _client = client;
            _stream = new MemoryStream();
            _path = path;
            existingStream.CopyTo(_stream);
        }

        public override bool CanRead => false;

        public override bool CanSeek => true;

        public override bool CanWrite => true;

        public override long Length => _stream.Length;

        public override long Position { get => _stream.Position ; set => _stream.Position = value; }

        public override void Flush()
        {
            _stream.Position = 0;
            _client.UploadOneDriveItemAsync(_path, _stream);
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _stream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            _stream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _stream.Write(buffer, offset, count);
        }
    }
}
