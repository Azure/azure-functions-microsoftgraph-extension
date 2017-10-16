// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.IO;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;

    internal class OneDriveStream : Stream
    {
        private readonly IGraphServiceClient _client;
        private readonly Stream _stream;
        private readonly string _path;
        private readonly FileAccess _fileAccess;

        public OneDriveStream(IGraphServiceClient client, FileAccess? fileAccess, Stream existingStream, string path)
        {
            _client = client;
            _fileAccess = fileAccess ?? FileAccess.ReadWrite;
            _path = path;
            _stream = CanWrite ? CopyStream(existingStream) : existingStream;
        }

        public override bool CanRead => _fileAccess == FileAccess.Read || _fileAccess == FileAccess.ReadWrite;

        public override bool CanSeek =>  true;

        public override bool CanWrite => _fileAccess == FileAccess.Write || _fileAccess == FileAccess.ReadWrite;

        public override long Length => _stream.Length;

        public override long Position { get => _stream.Position ; set => _stream.Position = value; }

        public override void Flush()
        {
            if (CanWrite)
            {
                var streamCopy = CopyStream(_stream);
                _client.UploadOneDriveItemAsync(_path, streamCopy);
            }

            _stream.Flush();         
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if(!CanRead)
            {
                throw new NotSupportedException("Cannot read from this stream without read permissions.");
            }
            return _stream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _stream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            if(!CanWrite)
            {
                throw new NotSupportedException("Cannot set the length without write permissions");
            }
            _stream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if(!CanWrite)
            {
                throw new NotSupportedException("Cannot write to this stream without write permissions");
            }
            _stream.Write(buffer, offset, count);
        }

        public override void Close()
        {
            this.Flush();
            base.Close();
            _stream.Close();
        }

        private static Stream CopyStream(Stream stream)
        {
            stream.Position = 0;
            var copyStream = new MemoryStream();
            stream.CopyTo(copyStream);
            copyStream.Position = 0;
            return copyStream;
        }
    }
}
