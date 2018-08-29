// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.Linq.Expressions;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Moq;

    internal static class OutlookMockUtilities
    {
        public static void MockSendMessage(this Mock<IGraphServiceClient> mock)
        {
            mock.Setup(client => client
                .Me
                .SendMail(It.IsAny<Message>(), true)
                .Request(null)
                .PostAsync(It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);
        }

        public static void VerifySendMessage(this Mock<IGraphServiceClient> mock, Expression<Func<Message, bool>> messageCondition)
        {
            // First verify PostAsync() called
            mock.Verify(client => client
                .Me
                .SendMail(It.IsAny<Message>(), true)
                .Request(null)
                .PostAsync(It.IsAny<CancellationToken>()));

            // Now verify that message condition holds for sent message
            mock.Verify(client => client
                .Me
                .SendMail(It.Is<Message>(messageCondition), true));
        }

        public static void VerifyDidNotSendMessage(this Mock<IGraphServiceClient> mock)
        {
            mock.Verify(client => client
                .Me
                .SendMail(It.IsAny<Message>(), true)
                .Request(null)
                .PostAsync(It.IsAny<CancellationToken>()), Times.Never());
        }
    }
}
