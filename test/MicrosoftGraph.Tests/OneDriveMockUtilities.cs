// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.IO;
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
                .GetAsync()).Returns(Task.FromResult(returnValue));
        } 

        public static void MockGetOneDriveContentStreamFromShareAsync(this Mock<IGraphServiceClient> mock, Stream returnValue)
        {
            mock.Setup(client => client
                .Shares[It.IsAny<string>()]
                .Root
                .Content
                .Request(null)
                .GetAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void VerifyGetOneDriveContentStreamFromShareAsync(this Mock<IGraphServiceClient> mock, string shareToken)
        {
            mock.Verify(client => client
                .Shares[shareToken]
                .Root
                .Content
                .Request(null)
                .GetAsync());
        }

    }
}
