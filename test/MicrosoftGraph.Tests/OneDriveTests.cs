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

            byte[] expected = GetContentAsBytes();
            Assert.True(expected.SequenceEqual(bytes));
            ResetState();
        }

        [Fact]
        public static async Task Input_ShareStream_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamFromShareAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.ShareStreamInput");

            byte[] expected = GetContentAsBytes();
            Assert.True(expected.SequenceEqual(bytes));
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
        public static async Task Input_TextReader_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());

            await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.TextReaderInput");

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

            await Assert.ThrowsAnyAsync<Exception>(async () => await CommonUtilities.ExecuteFunction<OneDriveInputs>(clientMock, "OneDriveInputs.StreamInputTryWrite"));

            ResetState();
        }

        [Fact]
        public static async Task Input_DriveItem_ReturnsExpectedValue()
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

            clientMock.VerifyUploadOneDriveItemAsync(normalPath, stream => ReadStreamBytes(stream).SequenceEqual(GetContentAsBytes()));
            ResetState();
        }

        [Fact]
        public static async Task Output_Stream_WriteOnly()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetOneDriveContentStreamAsync(GetContentAsStream());
            clientMock.MockUploadOneDriveItemAsync(null);

            await Assert.ThrowsAnyAsync<Exception>(async() => await CommonUtilities.ExecuteFunction<OneDriveOutputs>(clientMock, "OneDriveOutputs.WriteStreamTryRead"));

            ResetState();
        }

        [Fact]
        public static async Task Output_Bytes_NewFileUploadsToOneDrive()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockUploadOneDriveItemAsync(null);
            clientMock.MockExceptionForGetOneDriveContentStreamAsync(new Exception());

            await CommonUtilities.ExecuteFunction<OneDriveOutputs>(clientMock, "OneDriveOutputs.WriteBytes");

            clientMock.VerifyUploadOneDriveItemAsync(normalPath, stream => ReadStreamBytes(stream).SequenceEqual(ReadStreamBytes(GetContentAsStream())));
            ResetState();
        }

        [Fact]
        public static async Task Output_TextWriter_NewFileUploadsToOneDrive()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockUploadOneDriveItemAsync(null);
            clientMock.MockExceptionForGetOneDriveContentStreamAsync(new Exception());

            await CommonUtilities.ExecuteFunction<OneDriveOutputs>(clientMock, "OneDriveOutputs.WriteTextWriter");

            clientMock.VerifyUploadOneDriveItemAsync(normalPath, stream => ReadStreamBytes(stream).SequenceEqual(ReadStreamBytes(GetContentAsStream())));
            ResetState();
        }

        private static void ResetState()
        {
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
            readStream.Capacity = (int) readStream.Length;
            return readStream.GetBuffer();
        }

        public class OneDriveInputs
        {
            public static void StreamInput([OneDrive(normalPath, FileAccess.Read)] Stream input)
            {
                bytes = ReadStreamBytes(input);
            }

            public static void StreamInputTryWrite([OneDrive(normalPath, FileAccess.Read)] Stream input)
            {
                try
                {
                    input.Write(new byte[1], 0, 1);
                }
                catch (Exception ex) when(!(ex is NotSupportedException))
                { 
                    //swallow any other exceptions
                }

    }

            public static void ShareStreamInput([OneDrive(sharePath, FileAccess.Read)] Stream input)
            {
                bytes = ReadStreamBytes(input);
            }

            public static void StringInput([OneDrive(normalPath, FileAccess.Read)] string input)
            {
                stringValue = input;
            }

            public static void BytesInput([OneDrive(normalPath, FileAccess.Read)] byte[] input)
            {
                bytes = input;
            }

            public static void TextReaderInput([OneDrive(normalPath, FileAccess.Read)] TextReader input)
            {
                stringValue = input.ReadLine();
            }

            public static void DriveItemInput([OneDrive(normalPath, FileAccess.Read)] DriveItem input)
            {
                driveItem = input;
            }
        }

        public class OneDriveOutputs
        {
            public static void WriteStream([OneDrive(normalPath, FileAccess.Write)] Stream output)
            {
                byte[] bytes = GetContentAsBytes();
                output.Write(bytes, 0, bytes.Length);
            }

            public static void WriteStreamTryRead([OneDrive(normalPath, FileAccess.Write)] Stream output)
            {
                try
                {
                    output.Read(new byte[1], 0, 1);
                }
                catch (Exception ex) when (!(ex is NotSupportedException))
                {
                    //swallow any other exceptions
                }
            }

            public static void WriteBytes([OneDrive(normalPath, FileAccess.Write)] out byte[] bytes)
            {
                bytes = GetContentAsBytes();
            }

            public static void WriteTextWriter([OneDrive(normalPath, FileAccess.Write)] TextWriter writer)
            {
                writer.Write(GetContentAsString());
            }
        }

    }
}
