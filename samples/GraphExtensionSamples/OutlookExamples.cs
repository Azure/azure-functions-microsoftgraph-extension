// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
namespace GraphExtensionSamples
{
    using System.Collections.Generic;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    public static class OutlookExamples
    {
        //Sending messages

        public static void SendMailFromMessageObject([Outlook(
            Identity = TokenIdentityMode.UserFromRequest)] out Message message)
        {
            message = new Message();
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/src/Microsoft.Graph/Models/Generated/Message.cs
            //for usage
        }

        public static void SendMailFromJObject([Outlook(
            Identity = TokenIdentityMode.UserFromRequest)] out JObject message)
        {
            message = new JObject();
            //See https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/message for a json
            //representation. Note that a "recipient" or "recipients" field takes the place of the ToRecipient field in
            //the schema, with "recipient" being a single object and "recipients" being multiple
        }

        public static void SendMailFromPoco([Outlook(
            Identity = TokenIdentityMode.UserFromRequest)] out MessagePoco message)
        {
            message = new MessagePoco();
            //See https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/message for a json
            //representation. Note that a "recipient" or "recipients" field takes the place of the ToRecipient field in
            //the schema, with "recipient" being a single object and "recipients" being multiple
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
