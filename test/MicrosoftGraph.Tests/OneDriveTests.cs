// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Moq;
    using Xunit;

    public class OneDriveTests
    {
        private static Stream stream;
        private static string stringValue;
        private static byte[] bytes;
        private static DriveItem driveItem;

        private static Encoding encoding = Encoding.UTF8;
        private const string normalPath = "sample/path.txt";
        private const string sharePath = "https://microsoft-my.sharepoint.com/:t:/p/comcmaho/randomstringhere";
        private const string encodedSharePath = "u!aHR0cHM6Ly9taWNyb3NvZnQtbXkuc2hhcmVwb2ludC5jb20vOnQ6L3AvY29tY21haG8vcmFuZG9tc3RyaW5naGVyZQ";

        [Fact]
        public static async Task Input_Stream_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.StreamInput");

            Stream expected = GetContentAsStream();
            Assert.Equal(ReadStreamBytes(expected), ReadStreamBytes(stream));
            ResetState();
        }

        [Fact]
        public static async Task Input_ShareStream_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamFromShareAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.ShareStreamInput");

            Stream expected = GetContentAsStream();
            Assert.Equal(ReadStreamBytes(expected), ReadStreamBytes(stream));
            clientMock.VerifyGetOneDriveContentStreamFromShareAsync(encodedSharePath);

            ResetState();
        }

        [Fact]
        public static async Task Input_String_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.StringInput");

            string expected = GetContentAsString();
            Assert.Equal(expected, stringValue);
            ResetState();
        }

        [Fact]
        public static async Task Input_Bytes_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.BytesInput");

            byte[] expected = GetContentAsBytes();
            Assert.Equal(expected, bytes);
            ResetState();
        }

        [Fact]
        public static async Task Input_Stream_IsReadOnly()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.StreamInput");

            Assert.Equal(false, stream.CanWrite);
            Assert.Equal(true, stream.CanRead);
            Assert.Throws<NotSupportedException>(() => stream.Write(null, 0, 0));
            ResetState();
        }

        [Fact]
        public static async Task Input_DriveItem_IsReadOnly()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            var returnedDrive = new DriveItem();
            clientMock.MockGetOneDriveItemAsync(returnedDrive);

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.DriveItemInput");

            DriveItem expected = returnedDrive;
            Assert.Equal(expected, driveItem);
            ResetState();
        }

        [Fact]
        public static async Task Output_Stream_UploadsToOneDriveOnFlush()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());
            clientMock.MockUploadOneDriveItemAsync(null);

            await CommonUtilities.ExecuteFunction<OneDriveOutputs>(clientMock, "OneDriveOutputs.WriteStream");

            //Flush the stream to upload to one drive
            stream.Flush();

            clientMock.VerifyUploadOneDriveItemAsync(normalPath, stream => ReadStreamBytes(stream).SequenceEqual(ReadStreamBytes(GetContentAsStream())));
            ResetState();
        }

        [Fact]
        public static async Task Output_Stream_WriteOnly()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());
            clientMock.MockUploadOneDriveItemAsync(null);

            await CommonUtilities.ExecuteFunction<OneDriveOutputs>(clientMock, "OneDriveOutputs.WriteStream");

            Assert.Equal(false, stream.CanRead);
            Assert.Equal(true, stream.CanWrite);
            Assert.Throws<NotSupportedException>(() => stream.Read(null, 0, 0));
            ResetState();
        }

        [Fact]
        public static async Task Output_Bytes_UploadsToOneDrive()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockUploadOneDriveItemAsync(null);

            await CommonUtilities.ExecuteFunction<OneDriveOutputs>(clientMock, "OneDriveOutputs.WriteBytes");

            clientMock.VerifyUploadOneDriveItemAsync(normalPath, stream => ReadStreamBytes(stream).SequenceEqual(ReadStreamBytes(GetContentAsStream())));
            ResetState();
        }

        private static void ResetState()
        {
            stream = null;
            stringValue = null;
            bytes = null;
            driveItem = null;
        }

        private static string GetContentAsString()
        {
            return "stringContent";
        }

        private static byte[] GetContentAsBytes()
        {
            return encoding.GetBytes(GetContentAsString());
        }

        private static Stream GetContentAsStream()
        {
            return new MemoryStream(GetContentAsBytes());
        }


        private static byte[] ReadStreamBytes(Stream stream)
        {
            MemoryStream readStream = new MemoryStream();
            stream.CopyTo(readStream);
            return readStream.GetBuffer();
        }

        public class OneDriveInputs
        {
            public static void StreamInput([OneDrive(normalPath, FileAccess.Read)] Stream input)
            {
                stream = input;
            }

            public static void ShareStreamInput([OneDrive(Path = sharePath)] Stream input)
            {
                stream = input;
            }

            public static void StringInput([OneDrive(Path = normalPath)] string input)
            {
                stringValue = input;
            }

            public static void BytesInput([OneDrive(Path = normalPath)] byte[] input)
            {
                bytes = input;
            }

            public static void DriveItemInput([OneDrive(Path = normalPath)] DriveItem input)
            {
                driveItem = input;
            }
        }

        public class OneDriveOutputs
        {
            public static void WriteStream([OneDrive(normalPath, FileAccess.Write)] Stream output)
            {
                stream = output;
            }

            public static void WriteBytes([OneDrive(Path = normalPath)] out byte[] output)
            {
                output = GetContentAsBytes();
            }
        }

    }
}
