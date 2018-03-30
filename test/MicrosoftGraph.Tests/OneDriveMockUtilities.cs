// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.IO;
    using System.Linq.Expressions;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Moq;

    internal static class OneDriveMockUtilities
    {
        public static void MockGetOneDriveContentStreamAsync(this Mock<IGraphServiceClient> mock, Stream returnValue)
        {
            mock.Setup(client => client.Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Content
                .Request(null)
                .GetAsync())
                .Returns(Task.FromResult(returnValue));
        } 

        public static void MockGetOneDriveContentStreamFromShareAsync(this Mock<IGraphServiceClient> mock, Stream returnValue)
        {
            mock.Setup(client => client
                .Shares[It.IsAny<string>()]
                .Root
                .Content
                .Request(null)
                .GetAsync())
                .Returns(Task.FromResult(returnValue));
        }

        public static void VerifyGetOneDriveContentStreamFromShareAsync(this Mock<IGraphServiceClient> mock, string shareToken)
        {
            //First verify GetAsync() called
            mock.Verify(client => client
                .Shares[shareToken]
                .Root
                .Content
                .Request(null)
                .GetAsync());

            //Then verify sharetoken correct
            mock.Verify(client => client
                .Shares[shareToken]);
        }

        public static void MockGetOneDriveItemAsync(this Mock<IGraphServiceClient> mock, DriveItem returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Request()
                .GetAsync())
                .Returns(Task.FromResult(returnValue));
        }

        public static void MockUploadOneDriveItemAsync(this Mock<IGraphServiceClient> mock, DriveItem returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Content
                .Request(null)
                .PutAsync<DriveItem>(It.IsAny<Stream>()))
                .Returns(Task.FromResult(returnValue));
        }

        public static void VerifyUploadOneDriveItemAsync(this Mock<IGraphServiceClient> mock, string path, Expression<Func<Stream, bool>> streamCondition)
        {
            mock.Verify(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Content
                .Request(null)
                .PutAsync<DriveItem>(It.Is<Stream>(streamCondition)));
        }

    }
}
