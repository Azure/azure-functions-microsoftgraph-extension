// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

using Microsoft.Azure.WebJobs.Host;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TokenBinding
{
    public class EasyAuthClientFactory
    {
        public virtual IEasyAuthClient GetClient(string hostName, string signingKey, TraceWriter log)
        {
            return new EasyAuthTokenClient(hostName, signingKey, log);
        }
    }
}
