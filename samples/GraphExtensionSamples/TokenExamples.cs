// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
namespace GraphExtensionSamples
{
    using System;
    using Microsoft.Azure.WebJobs;

    public static class TokenExamples
    {
        [NoAutomaticTrigger]
        public static void GraphTokenFromId([Token(
                UserId = "%UserId%", 
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Resource = "https://graph.microsoft.com")] string token)
        {
            Console.Write("The Microsoft graph token for the user is: " + token);
        }

        [NoAutomaticTrigger]
        public static void GraphTokenFromUserToken([Token(
                Identity = TokenIdentityMode.UserFromToken,
                UserToken = "%UserToken%",
                IdentityProvider = "AAD",
                Resource = "https://graph.microsoft.com")] string token)
        {
            Console.Write("The microsoft graph token for the user is: " + token);
        }

        // NOTE: This would only work in a Function with an HTTP trigger with
        // requests having the header X-MS-TOKEN-AAD-ID-TOKEN
        [NoAutomaticTrigger]
        public static void GraphTokenFromHttpRequest([Token(
                Identity = TokenIdentityMode.UserFromRequest,
                IdentityProvider = "AAD",
                Resource = "https://graph.microsoft.com")] string token)
        {
            Console.Write("The microsoft graph token for the user is: " + token);
        }

        // This template uses application permissions and requires consent from an Azure Active Directory admin.
        // See https://go.microsoft.com/fwlink/?linkid=858780
        [NoAutomaticTrigger]
        public static void GraphTokenFromApplicationIdentity([Token(
                Identity = TokenIdentityMode.ClientCredentials,
                IdentityProvider = "AAD",
                Resource = "https://graph.microsoft.com")] string token)
        {
            Console.Write("The microsoft graph token for the application is: " + token);
        }

    }
}
