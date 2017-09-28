// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.IO;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Moq;
    using Xunit;

    public class OneDriveTests
    {
        private static Stream stream;

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

        private static void ResetState()
        {
            stream = null;
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
            public static void StreamInput([OneDrive(Path = normalPath)] Stream input)
            {
                stream = input;
            }

            public static void ShareStreamInput([OneDrive(Path = sharePath)] Stream input)
            {
                stream = input;
            }
        }

    }
}
