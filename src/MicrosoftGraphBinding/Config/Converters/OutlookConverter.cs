// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    internal class OutlookConverter : IConverter<JObject, Message>, IConverter<string, Message>
    {
        public Message Convert(JObject input)
        {
            // Set up recipient(s)
            List<Recipient> recipientList = new List<Recipient>();

            JToken recipientToken = GetPropertyValueIgnoreCase<JToken>(input, "recipient", throwException: false)
                ?? GetPropertyValueIgnoreCase<JToken>(input, "recipients", throwException: false);

            if (recipientToken == null)
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
                        Name = GetPropertyValueIgnoreCase<string>(recip, "name", throwException: false),
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

        public Message Convert(string input)
        {
            return Convert(JObject.Parse(input));
        }

        private static T GetPropertyValueIgnoreCase<T>(JObject input, string key, bool throwException = true)
        {
            JToken value;
            if (!input.TryGetValue(key, StringComparison.OrdinalIgnoreCase, out value))
            {
                if (throwException)
                {
                    throw new InvalidOperationException($"The object needs to have a {key} field.");
                }
                return default(T);
            }
            return value.ToObject<T>();
        }
    }

    // This converter goes directly to Message instead of T -> JObject and composing 
    // with JObject -> Message as composition conversions with OpenTypes are broken
    internal class OutlookGenericsConverter<T> : IConverter<T, Message>
    {
        private static readonly OutlookConverter Converter = new OutlookConverter(); 

        public Message Convert(T input)
        {
            return Converter.Convert(JObject.FromObject(input));
        }
    }
}
