// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;
using Microsoft.Azure.WebJobs.Extensions.AuthTokens;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    public class GraphOptions : TokenOptions
    {
        public string SubscriptionStoreLocationAppSettingName { get; set; } = "BYOB_TokenMap";

        public string GraphBaseUrl { get; set; } = O365Constants.GraphBaseUrl;

        public string TokenMapLocation { get; set; } = "D:/home/data/byob_graphmap";

        public TimeSpan WebhookExpirationTimeSpan { get; set; } = new TimeSpan(0, 0, 4230, 0);

        public override void SetAppSettings(INameResolver appSettings)
        {
            if (TokenMapLocation == null)
            {
                TokenMapLocation = appSettings.Resolve(SubscriptionStoreLocationAppSettingName);
            }
            base.SetAppSettings(appSettings);
        }
    }
}
