// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using System;

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config
{
    public class GraphOptions
    {
        private const string defaultTokenMapLocation = "D:/home/data/byob_graphmap";
        public string SubscriptionStoreLocationAppSettingName { get; set; } = "BYOB_TokenMap";

        public string GraphBaseUrl { get; set; } = O365Constants.GraphBaseUrl;

        public string TokenMapLocation { get; set; } = defaultTokenMapLocation;

        public TimeSpan WebhookExpirationTimeSpan { get; set; } = new TimeSpan(0, 0, 4230, 0);

        public void SetAppSettings(INameResolver appSettings)
        {
            var settingsTokenMapLocation = appSettings.Resolve(SubscriptionStoreLocationAppSettingName);
            if (!string.IsNullOrEmpty(settingsTokenMapLocation) && TokenMapLocation == defaultTokenMapLocation)
                TokenMapLocation = settingsTokenMapLocation;
        }
    }
}
