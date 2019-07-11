using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    public interface IAadServiceFactory
    {
        IAadService GetAadClient(string tenantUrl, string clientId, string clientSecret);
    }
}
