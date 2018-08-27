// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph;

    internal static class OutlookClient
    {
        /// <summary>
        /// Send an email with a dynamically set body
        /// </summary>
        /// <param name="client">GraphServiceClient used to send request</param>
        /// <returns>Async task for posted message</returns>
        public static async Task SendMessageAsync(this IGraphServiceClient client, Message msg, CancellationToken token)
        {
            // Send message & save to sent items folder
            await client.Me.SendMail(msg, true).Request().PostAsync(token);
        }
    }
}
