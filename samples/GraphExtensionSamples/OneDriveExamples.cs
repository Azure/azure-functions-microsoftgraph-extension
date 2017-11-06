// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace GraphExtensionSamples
{
    using System;
    using System.IO;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;

    public static class OneDriveExamples
    {
        //NOTE: In the current release, if these bindings are being used in Azure Functions, you
        //should explicitly set a property Access to Read or Write, depending on which operation
        //you are trying to perform. In the future, this will be autopopulated based on whether it
        //is an input or an output binding.

        public static void ReadOneDriveFileAsByteArray([OneDrive(
            Identity = TokenIdentityMode.UserFromId,
            UserId = "sampleuserid",
            Path = "samplepath.txt")] byte[] array)
        {
            Console.Write(System.Text.Encoding.UTF8.GetString(array, 0, array.Length));
        }


        //NOTE: These strings read the file assuming UTF-8 encoding.
        public static void ReadOneDriveFileAsString([OneDrive(
            Identity = TokenIdentityMode.UserFromId,
            UserId = "sampleuserid",
            Path = "samplepath.txt")] string fileText)
        {
            Console.Write(fileText);
        }

        //The binding also supports paths in the form of share links
        public static void ReadOneDriveFileAsByteArrayFromShareLink([OneDrive(
            Identity = TokenIdentityMode.UserFromId,
            UserId = "sampleuserid",
            Path = "https://microsoft-my.sharepoint.com/:t:/p/comcmaho/randomstringhere")] byte[] array)
        {
            Console.Write(System.Text.Encoding.UTF8.GetString(array, 0, array.Length));
        }

        public static void GetDriveItem([OneDrive(
            Identity = TokenIdentityMode.UserFromId,
            UserId = "sampleuserid",
            Path = "samplepath.txt")] DriveItem array)
        {
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/81c50e72166152f9f84dc38b2516379b7a536300/src/Microsoft.Graph/Models/Generated/DriveItem.cs
            //for usage
        }

        public static void GetOneDriveStream([OneDrive(
            Identity = TokenIdentityMode.UserFromId,
            UserId = "sampleuserid",
            Path = "samplepath.txt")] Stream stream)
        {
            byte[] buffer = new byte[256];
            stream.Read(buffer, 0, 256);
        }

        public static void GetOneDriveStreamWithWriteAccess([OneDrive(FileAccess.Write, 
            Identity = TokenIdentityMode.UserFromId,
            UserId = "sampleuserid",
            Path = "samplepath.txt")] Stream stream)
        {
            string sampleText = "sampleText";
            byte[] encodedText = System.Text.Encoding.UTF8.GetBytes(sampleText);
            stream.Write(encodedText, 0, encodedText.Length);
        }
    }
}
