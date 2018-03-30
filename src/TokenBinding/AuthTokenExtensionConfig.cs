// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.Token.Tests")]
[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests")]
namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
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
        // Useful for binding to additional inputs
        private FluentBindingRule<TokenAttribute> TokenRule { get; set; }

        internal IEasyAuthClient EasyAuthClient
        {
            get
            {
                if (_easyAuthClient == null)
                {
                    string hostname = AppSettings.Resolve(Constants.AppSettingWebsiteHostname);
                    _easyAuthClient = new EasyAuthTokenClient(hostname, LoggerFactory);
                }
                return _easyAuthClient;
            }

            set
            {
                _easyAuthClient = value;
            }
        }

        internal IAadClient AadClient
        {
            get
            {
                if (_aadManager == null)
                {
                    string clientId = AppSettings.Resolve(Constants.AppSettingClientIdName);
                    string clientSecret = AppSettings.Resolve(Constants.AppSettingClientSecretName);
                    string tenantUrl = AppSettings.Resolve(Constants.AppSettingWebsiteAuthOpenIdIssuer);
                    _aadManager = new AadClient(new ClientCredential(clientId, clientSecret), tenantUrl);
                }
                return _aadManager;
            }
            set
            {
                _aadManager = value;
            }
        }

        internal ILoggerFactory LoggerFactory;

        internal INameResolver AppSettings { get; set; }

        private IAadClient _aadManager;

        private IEasyAuthClient _easyAuthClient;


        //TODO: https://github.com/Azure/azure-functions-microsoftgraph-extension/issues/48
        internal static string CreateBindingCategory(string bindingName)
        {
            return $"Host.Bindings.{bindingName}";
        }

        /// <summary>
        /// Retrieve audience from raw JWT
        /// </summary>
        /// <param name="rawToken">JWT</param>
        /// <returns>Token audience</returns>
        private static string GetAudience(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var audience = jwt.Audiences.FirstOrDefault();
            return audience;
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
        }

        public void Initialize(ExtensionConfigContext context)
        {
            InitializeAllExceptRules(context);
            var converter = new Converters(this);
            this.TokenRule = context.AddBindingRule<TokenAttribute>();
            this.TokenRule.BindToInput<string>(converter);
        }

        /// <summary>
        /// Retrieve an access token for the specified resource (e.g. MS Graph)
        /// </summary>
        /// <param name="attribute">TokenAttribute with desired resource & user's principal ID or ID token</param>
        /// <returns>JWT with audience, scopes, user id</returns>
        public async Task<string> GetAccessTokenAsync(TokenAttribute attribute)
        {
            attribute.CheckValidity();
            switch (attribute.Identity)
            {
                case TokenIdentityMode.UserFromId:
                    // If the attribute has no identity provider, assume AAD
                    attribute.IdentityProvider = attribute.IdentityProvider ?? "AAD";
                    string signingKey = AppSettings.Resolve(Constants.AppSettingWebsiteAuthSigningKey);
                    var easyAuthTokenManager = new EasyAuthTokenManager(EasyAuthClient, signingKey);
                    return await easyAuthTokenManager.GetEasyAuthAccessTokenAsync(attribute);
                case TokenIdentityMode.UserFromToken:
                    return await GetAuthTokenFromUserToken(attribute.UserToken, attribute.Resource);
                case TokenIdentityMode.ClientCredentials:
                    return await AadClient.GetTokenFromClientCredentials(attribute.Resource);
            }

            throw new InvalidOperationException("Unable to authorize without Principal ID or ID Token.");
        }

        private async Task<string> GetAuthTokenFromUserToken(string userToken, string resource)
        {
            if (string.IsNullOrWhiteSpace(resource))
            {
                throw new ArgumentException("A resource is required to get an auth token on behalf of a user.");
            }

            // If the incoming token already has the correct audience (resource), then skip the exchange (it will fail with AADSTS50013!)
            var currentAudience = GetAudience(userToken);
            if (currentAudience != resource)
            {
                string token = await AadClient.GetTokenOnBehalfOfUserAsync(
                    userToken,
                    resource);
                return token;
            }

            // No exchange requested, return token directly.
            return userToken;
        }



        /// <summary>
        /// Converter class used to convert TokenAttribute from binding -> different kinds of inputs to fx
        /// </summary>
        public class Converters :
            IAsyncConverter<TokenAttribute, string>
        {
            private readonly AuthTokenExtensionConfig _parent;

            /// <summary>
            /// Initializes a new instance of the <see cref="Converters"/> class.
            /// </summary>
            /// <param name="parent">TokenExtensionConfig containing necessary context & methods</param>
            public Converters(AuthTokenExtensionConfig parent)
            {
                this._parent = parent;
            }

            /// <summary>
            /// Convert from a TokenAttribute to a string (to use as input to a fx)
            /// </summary>
            /// <param name="input">TokenAttribute with necessary user info & desired resource</param>
            /// <param name="cancellationToken">Used to propagate notifications</param>
            /// <returns>JWT</returns>
            async Task<string> IAsyncConverter<TokenAttribute, string>.ConvertAsync(TokenAttribute input, CancellationToken cancellationToken)
            {
                var token = await this._parent.GetAccessTokenAsync(input);
                return token;
            }
        }
    }
}
