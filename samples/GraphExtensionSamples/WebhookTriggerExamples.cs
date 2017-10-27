// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace GraphExtensionSamples
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;

    public static class WebhookTriggerExamples
    {
        //NOTE: There can only be one trigger per resource type. That means all graph webhook subscriptions for the 
        //given resource type will go through the same function.

        public static void OnMessage([GraphWebhookTrigger(ResourceType = "#Microsoft.Graph.Message")] Message message)
        {
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/src/Microsoft.Graph/Models/Generated/Message.cs
            //for usage
        }

        public static void OnOneDriveChange([GraphWebhookTrigger(ResourceType = "#Microsoft.Graph.DriveItem")] DriveItem driveItem)
        {
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/src/Microsoft.Graph/Models/Generated/DriveItem.cs
            //for usage
        }

        public static void OnContactChange([GraphWebhookTrigger(ResourceType = "#Microsoft.Graph.Contact")] Contact contact)
        {
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/src/Microsoft.Graph/Models/Generated/Contact.cs
            //for usage
        }

        public static void OnEventChange([GraphWebhookTrigger(ResourceType = "#Microsoft.Graph.Event")] Event eventNotice) 
        {
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/src/Microsoft.Graph/Models/Generated/Event.cs
            //for usage
        }
        }
}
