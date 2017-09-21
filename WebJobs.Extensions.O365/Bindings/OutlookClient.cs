// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;

    internal static class OutlookClient
    {
        /// <summary>
        /// Using the JTokens of a JObject, construct a Microsoft.Graph.Message object
        /// </summary>
        /// <param name="input">JObject with subject, body, and recipient(s)</param>
        /// <returns>Microsoft.Graph.Message object used by MS Graph Service Client to send email</returns>
        public static Message CreateMessage(JObject input)
        {
            // Set up recipient(s)
            List<Recipient> recipientList = new List<Recipient>();

            var r = input["recipient"] ?? input["recipients"]; // Grab either single recipient JObject or JArray of recipients

            List<JObject> recipients;

            // MS Graph Message expects a list of recipients
            if (r is JArray)
            {
                // JArray -> List
                recipients = r.ToObject<List<JObject>>();
            }
            else
            {
                // List with one JObject
                recipients = new List<JObject>();
                recipients.Add(r.ToObject<JObject>());
            }

            if (recipients.Count == 0)
            {
                throw new InvalidOperationException("At least one recipient must be provided.");
            }

            foreach (JObject recip in recipients)
            {
                var name = recip["name"]?.ToString();
                Recipient recipient = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recip["address"].ToString(),
                        Name = name,
                    },
                };
                recipientList.Add(recipient);
            }

            // Actually create message
            var msg = new Message
            {
                Body = new ItemBody
                {
                    Content = input["body"].ToString(),
                    ContentType = BodyType.Text,
                },
                Subject = input["subject"].ToString(),
                ToRecipients = recipientList,
            };

            return msg;
        }

        /// <summary>
        /// Parse json string then use main converter class to convert to Message object
        /// </summary>
        /// <param name="msg">JSON formatted string</param>
        /// <returns>Message object to be sent</returns>
        public static Message CreateMessage(string msg)
        {
            return CreateMessage(JObject.Parse(msg));
        }

        /// <summary>
        /// Send an email with a dynamically set body
        /// </summary>
        /// <param name="client">GraphServiceClient used to send request</param>
        /// <param name="attr">Outlook Attribute with necessary data to build request</param>
        /// <returns>Async task for posted message</returns>
        public static async Task SendMessage(this GraphServiceClient client, OutlookAttribute attr, Message msg)
        {
            // Send message & save to sent items folder
            await client.Me.SendMail(msg, true).Request().PostAsync();
        }
    }
}
