// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.Token.Tests")]
[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests")]
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
