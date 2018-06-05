// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.Token.Tests")]
[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests")]
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.Extensions.Logging;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    /// <summary>
    /// WebJobs SDK Extension for Token binding.
    /// </summary>
    public class AuthTokenExtensionConfig : IExtensionConfigProvider
    {
        internal IEasyAuthClient EasyAuthClient { get; set; }

        internal IAadClient AadClient { get; set; }

        internal ILoggerFactory LoggerFactory { get; set; }

        internal INameResolver AppSettings { get; set; }

        private IAsyncConverter<TokenAttribute, string> _converter;


        //TODO: https://github.com/Azure/azure-functions-microsoftgraph-extension/issues/48
        internal static string CreateBindingCategory(string bindingName)
        {
            return $"Host.Bindings.{bindingName}";
        }

        /// <summary>
        /// Initialize the binding extension
        /// </summary>
        /// <param name="context">Context for extension</param>
        public void InitializeAllExceptRules(ExtensionConfigContext context)
        {
            var config = context.Config;
            // Set up logging
            LoggerFactory = context.Config.LoggerFactory ?? throw new ArgumentNullException("No logger present");
            AppSettings = AppSettings ?? config.NameResolver;
            EasyAuthClient = EasyAuthClient ?? new EasyAuthTokenClient(AppSettings, LoggerFactory);
            AadClient = AadClient ?? new AadClient(AppSettings);
        }

        public void Initialize(ExtensionConfigContext context)
        {
            InitializeAllExceptRules(context);
            _converter = new TokenConverter(EasyAuthClient, AadClient, AppSettings);
            var tokenRule = context.AddBindingRule<TokenAttribute>();
            tokenRule.BindToInput<string>(_converter);
        }

        /// <summary>
        /// Retrieve an access token for the specified resource (e.g. MS Graph)
        /// </summary>
        /// <param name="attribute">TokenAttribute with desired resource & user's principal ID or ID token</param>
        /// <returns>JWT with audience, scopes, user id</returns>
        public async Task<string> GetAccessTokenAsync(TokenAttribute attribute)
        {
            return await _converter.ConvertAsync(attribute, CancellationToken.None);
        }
    }
}
