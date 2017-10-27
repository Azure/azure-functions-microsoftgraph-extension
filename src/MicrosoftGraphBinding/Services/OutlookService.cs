// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    internal class OutlookService
    {
        private IGraphServiceClient _client;

        public OutlookService(IGraphServiceClient client)
        {
            _client = client;
        }

        public async Task SendMessageAsync(Message msg)
        {
            await _client.SendMessageAsync(msg);
        }

        public static T GetPropertyValueIgnoreCase<T>(JObject input, string key, bool throwException = true)
        {
            JToken value;
            if(!input.TryGetValue(key, StringComparison.OrdinalIgnoreCase, out value)) {
                if (throwException)
                {
                    throw new InvalidOperationException($"The object needs to have a {key} field.");
                }
                return default(T);
            }
            return value.ToObject<T>();
        }

        /// <summary>
        /// Using the JTokens of a JObject, construct a Microsoft.Graph.Message object
        /// </summary>
        /// <param name="input">JObject with subject, body, and recipient(s)</param>
        /// <returns>Microsoft.Graph.Message object used by MS Graph Service Client to send email</returns>
        public static Message CreateMessage(JObject input)
        {
            // Set up recipient(s)
            List<Recipient> recipientList = new List<Recipient>();

            JToken recipientToken = GetPropertyValueIgnoreCase<JToken>(input, "recipient", false) 
                ?? GetPropertyValueIgnoreCase<JToken>(input, "recipients", false);

            if(recipientToken == null)
            {
                throw new InvalidOperationException("The object needs to have a 'recipient' or 'recipients' field.");
            }
                    
            List<JObject> recipients;

            // MS Graph Message expects a list of recipients
            if (recipientToken is JArray)
            {
                // JArray -> List
                recipients = recipientToken.ToObject<List<JObject>>();
            }
            else
            {
                // List with one JObject
                recipients = new List<JObject>();
                recipients.Add(recipientToken.ToObject<JObject>());
            }

            if (recipients.Count == 0)
            {
                throw new InvalidOperationException("At least one recipient must be provided.");
            }

            foreach (JObject recip in recipients)
            {
                Recipient recipient = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = GetPropertyValueIgnoreCase<string>(recip, "address"),
                        Name = GetPropertyValueIgnoreCase<string>(recip, "name", false),
                    },
                };
                recipientList.Add(recipient);
            }

            // Actually create message
            var msg = new Message
            {
                Body = new ItemBody
                {
                    Content = GetPropertyValueIgnoreCase<string>(input, "body"),
                    ContentType = BodyType.Text,
                },
                Subject = GetPropertyValueIgnoreCase<string>(input, "subject"),
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
    }
}
