// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    using Microsoft.Graph;

    internal interface IGraphServiceClientProvider
    {
        IGraphServiceClient CreateNewGraphServiceClient(string token);

        void UpdateGraphServiceClientAuthToken(IGraphServiceClient client, string token);
    }
}
