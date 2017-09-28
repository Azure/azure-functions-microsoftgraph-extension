// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Azure.WebJobs.Extensions.Token.Tests;
    using Microsoft.Graph;
    using Moq;
    using Newtonsoft.Json.Linq;
    using Xunit;

    public class OutlookTests
    {
        private static string name = "Sample User";
        private static string address = "sampleuser@microsoft.com";
        private static string subject = "hello";
        private static string body = "world";
        private static BodyType contentType = BodyType.Text;

        [Fact]
        public static async Task JObjectOutput_SendsProperMessage()
        {
            var clientMock = SendMessageMock();

            await CommonUtilities.ExecuteFunction<OutlookFunctions>(clientMock, "OutlookFunctions.SendJObject");

            clientMock.VerifySendMessage(msg => MessageEquals(msg, GetMessage()));
        }

        [Fact]
        public static async Task MessageOutput_SendsProperMessage()
        {
            var clientMock = SendMessageMock();

            await CommonUtilities.ExecuteFunction<OutlookFunctions>(clientMock, "OutlookFunctions.SendMessage");

            clientMock.VerifySendMessage(msg => MessageEquals(msg, GetMessage()));
        }

        [Fact]
        public static async Task PocoOutput_SendsProperMessage()
        {
            var clientMock = SendMessageMock();

            await CommonUtilities.ExecuteFunction<OutlookFunctions>(clientMock, "OutlookFunctions.SendPoco");

            clientMock.VerifySendMessage(msg => MessageEquals(msg, GetMessage()));
        }

        [Fact]
        public static async Task PocoOutputWithNoRecipients_ThrowsException()
        {
            var clientMock = SendMessageMock();

            await Assert.ThrowsAnyAsync<Exception>(async () => await CommonUtilities.ExecuteFunction<OutlookFunctions>(clientMock, "OutlookFunctions.NoRecipients"));
            clientMock.VerifyDidNotSendMessage();
        }

        private static Mock<IGraphServiceClient> SendMessageMock()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockSendMessage();
            return clientMock;
        }

        private static bool MessageEquals(Message message1, Message message2)
        {
            return message1.Body.Content.Equals(message2.Body.Content) &&
                message1.Body.ContentType.Equals(message2.Body.ContentType) &&
                message1.Subject.Equals(message2.Subject) &&
                RecipientsEqual(message1.ToRecipients, message2.ToRecipients);
        }

        private static bool RecipientsEqual(IEnumerable<Recipient> recipients1, IEnumerable<Recipient> recipients2)
        {
            return recipients1.Select(msg => msg.EmailAddress.Name).SequenceEqual(recipients2.Select(msg => msg.EmailAddress.Name)) &&
                recipients1.Select(msg => msg.EmailAddress.Address).SequenceEqual(recipients2.Select(msg => msg.EmailAddress.Address));
        }

        private static Message GetMessage()
        {
            return new Message()
            {
                Body = new ItemBody()
                {
                    Content = body,
                    ContentType = contentType,
                },
                Subject = subject,
                ToRecipients = new List<Recipient>()
                {
                    new Recipient()
                    {
                        EmailAddress = new EmailAddress()
                        {
                            Name = name,
                            Address = address,
                        }
                    }
                }
            };
        }

        private static JObject GetMessageAsJObject()
        {
            var message = new JObject();
            message["subject"] = subject;
            message["body"] = body;
            message["recipient"] = new JObject();
            message["recipient"]["address"] = address;
            message["recipient"]["name"] = name;
            return message;
        }

        private static MessagePoco GetMessageAsPoco()
        {
            return new MessagePoco()
            {
                Body = body,
                Subject = subject,
                Recipients = new List<RecipientPoco>()
                {
                    new RecipientPoco()
                    {
                        Address = address,
                        Name = name,
                    }
                }
            };
        }

        private class OutlookFunctions
        {
            public static void SendJObject([Outlook] out JObject message)
            {
                message = GetMessageAsJObject();
            }

            public static void SendMessage([Outlook] out Message message)
            {
                message = GetMessage();
            }

            public static void SendPoco([Outlook] out MessagePoco message)
            {
                message = GetMessageAsPoco();
            }

            public static void NoRecipients([Outlook] out MessagePoco message)
            {
                message = GetMessageAsPoco();
                message.Recipients = new List<RecipientPoco>();
            }
        }

        public class MessagePoco
        {
            public string Subject { get; set; }
            public string Body { get; set; }
            public List<RecipientPoco> Recipients { get; set; }
        }

        public class RecipientPoco
        {
            public string Address { get; set; }
            public string Name { get; set; }
        }
    }
}
