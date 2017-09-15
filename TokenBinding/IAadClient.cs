﻿// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace TokenBinding
{
    using System.Threading.Tasks;

    public interface IAadClient
    {
        Task<string> GetTokenOnBehalfOfUserAsync(string userToken, string resource);

        Task<string> GetTokenFromClientCredentials(string resource);
    }
}
