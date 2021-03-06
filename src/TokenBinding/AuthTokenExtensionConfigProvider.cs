﻿// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.Token.Tests, PublicKey=0024000004800000940000000602000000240000525341310004000001000100cd1dabd5a893b40e75dc901fe7293db4a3caf9cd4d3e3ed6178d49cd476969abe74a9e0b7f4a0bb15edca48758155d35a4f05e6e852fff1b319d103b39ba04acbadd278c2753627c95e1f6f6582425374b92f51cca3deb0d2aab9de3ecda7753900a31f70a236f163006beefffe282888f85e3c76d1205ec7dfef7fa472a17b1")]
[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests, PublicKey=0024000004800000940000000602000000240000525341310004000001000100cd1dabd5a893b40e75dc901fe7293db4a3caf9cd4d3e3ed6178d49cd476969abe74a9e0b7f4a0bb15edca48758155d35a4f05e6e852fff1b319d103b39ba04acbadd278c2753627c95e1f6f6582425374b92f51cca3deb0d2aab9de3ecda7753900a31f70a236f163006beefffe282888f85e3c76d1205ec7dfef7fa472a17b1")]
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2, PublicKey=0024000004800000940000000602000000240000525341310004000001000100cd1dabd5a893b40e75dc901fe7293db4a3caf9cd4d3e3ed6178d49cd476969abe74a9e0b7f4a0bb15edca48758155d35a4f05e6e852fff1b319d103b39ba04acbadd278c2753627c95e1f6f6582425374b92f51cca3deb0d2aab9de3ecda7753900a31f70a236f163006beefffe282888f85e3c76d1205ec7dfef7fa472a17b1")]
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// WebJobs SDK Extension for Token binding.
    /// </summary>
    internal class AuthTokenExtensionConfigProvider : IExtensionConfigProvider
    {
        private TokenConverter _converter;
        private TokenOptions _options;

        public AuthTokenExtensionConfigProvider(IOptions<TokenOptions> options, IAadClient aadClient, IEasyAuthClient easyAuthClient, INameResolver appSettings)
        {
            _options = options.Value;
            _options.SetAppSettings(appSettings);
            _converter = new TokenConverter(options, easyAuthClient, aadClient);
        }

        //TODO: https://github.com/Azure/azure-functions-microsoftgraph-extension/issues/48
        internal static string CreateBindingCategory(string bindingName)
        {
            return $"Host.Bindings.{bindingName}";
        }

        /// <summary>
        /// Initialize the binding extension
        /// </summary>
        /// <param name="context">Context for extension</param>
        public void Initialize(ExtensionConfigContext context)
        {
            var tokenRule = context.AddBindingRule<TokenAttribute>();
            tokenRule.BindToInput<string>(_converter);
        }
    }
}
