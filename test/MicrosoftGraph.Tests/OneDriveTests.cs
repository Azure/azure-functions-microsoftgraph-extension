// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Extensions.Token.Tests;
    using Microsoft.Graph;
    using Moq;
    using Xunit;

    public class OneDriveTests
    {
        private static string content;
        private static byte[] bytes;
        private static Stream stream;
        private static DriveItem driveItem;

        private static Encoding encoding = Encoding.UTF8;
        private const string normalPath = "sample/path.txt";
        private const string sharePath = "https://microsoft-my.sharepoint.com/:t:/p/comcmaho/randomstringhere";

        [Fact]
        public static async Task Input_String_ReturnsExpectedValue()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var oneDriveMock = new Mock<IOneDriveClient>();
            oneDriveMock.Setup(client => client.GetOneDriveContentStreamAsync(It.IsAny<string>())).Returns(Task.FromResult(GetContentAsStream()));
            graphConfig._onedriveClient = oneDriveMock.Object;

            var jobHost = TestHelpers.NewHost<OneDriveInputs>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("OneDriveInputs.StringInput", args);
            string expected = GetContentAsString();
            Assert.Equal(expected, content);
            ResetState();
        }

        private static void ResetState()
        {
            content = null;
            bytes = null;
            stream = null;
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

        private static DriveItem GetDriveItem()
        {
            return new DriveItem()
            {
                Content = GetContentAsStream(),
            };
        }

        private static byte[] ReadStreamBytes(Stream stream)
        {
            MemoryStream readStream = new MemoryStream();
            stream.CopyTo(readStream);
            return readStream.GetBuffer();
        }

        private static string ReadStreamText(Stream stream)
        {
            return encoding.GetString(ReadStreamBytes(stream));
        }


        public class OneDriveInputs
        {
            public static void StringInput([OneDrive(Path = normalPath)] string input)
            {
                content = input;
            }

            public static void BytesInput([OneDrive(Path = normalPath)] byte[] input)
            {
                bytes = input;
            }

            public static void StreamInput([OneDrive(Path = normalPath)] Stream input)
            {
                stream = input;
            }

            public static void DriveItemInput([OneDrive(Path = normalPath)] DriveItem input)
            {
                driveItem = input;
            }
        }

    }
}
