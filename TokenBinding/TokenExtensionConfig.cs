// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace TokenBinding
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Linq;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Host;
    using Microsoft.Azure.WebJobs.Host.Config;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;

    /// <summary>
    /// WebJobs SDK Extension for Token binding.
    /// </summary>
    public class TokenExtensionConfig : IExtensionConfigProvider
    {
        // Useful for binding to additional inputs
        public FluentBindingRule<TokenAttribute> TokenRule { get; set; }

        public EasyAuthClientFactory EasyAuthClientFactory {
            get
            {
                return _easyAuthClientFactory;
            }

            set
            {
                _easyAuthClientFactory = value;
            }
        }

        public AadClientFactory AadClientFactory
        {
            get
            {
                return _aadClientFactory;
            }

            set
            {
                _aadClientFactory = value;
            }
        }

        public INameResolver AppSettings { get; set; }

        private EasyAuthClientFactory _easyAuthClientFactory = new EasyAuthClientFactory();

        private AadClientFactory _aadClientFactory = new AadClientFactory();

        internal TraceWriter _log;

        private IAadClient _aadManager;

        /// <summary>
        /// Retrieve audience from raw JWT
        /// </summary>
        /// <param name="rawToken">JWT</param>
        /// <returns>Token audience</returns>
        public static string GetAudience(string rawToken)
        {
            var jwt = new JwtSecurityToken(rawToken);
            var audience = jwt.Audiences.FirstOrDefault();
            return audience;
        }

        public IAadClient GetAadClient()
        {
            if (_aadManager == null)
            {
                string clientId = AppSettings.Resolve(Constants.AppSettingClientIdName);
                string clientSecret = AppSettings.Resolve(Constants.AppSettingClientSecretName);
                _aadManager = AadClientFactory.GetClient(new ClientCredential(clientId, clientSecret));
            }
            return _aadManager;
        }

        private IEasyAuthClient GetEasyAuthTokenClient()
        {
            string hostname = AppSettings.Resolve(Constants.AppSettingWebsiteHostname);
            string signingKey = AppSettings.Resolve(Constants.AppSettingWebsiteAuthSigningKey);
            return EasyAuthClientFactory.GetClient(hostname, signingKey, _log);
        }

        /// <summary>
        /// Initialize the binding extension
        /// </summary>
        /// <param name="context">Context for extension</param>
        public void Initialize(ExtensionConfigContext context)
        {
            var config = context.Config;

            // Set up logging
            _log = context.Trace;

            AppSettings = AppSettings ?? config.NameResolver;

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
                case IdentityMode.UserFromId:
                    // If the attribute has no identity provider, assume AAD
                    attribute.IdentityProvider = attribute.IdentityProvider ?? "AAD";
                    var easyAuthTokenManager = new EasyAuthTokenManager(GetEasyAuthTokenClient());
                    return await easyAuthTokenManager.GetEasyAuthAccessTokenAsync(attribute);
                case IdentityMode.UserFromToken:
                    return await GetAuthTokenFromUserToken(attribute.UserToken, attribute.Resource);
                case IdentityMode.ClientCredentials:
                    return await GetAadClient().GetTokenFromClientCredentials(attribute.Resource);
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
                string token = await GetAadClient().GetTokenOnBehalfOfUserAsync(
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
            private readonly TokenExtensionConfig _parent;

            /// <summary>
            /// Initializes a new instance of the <see cref="Converters"/> class.
            /// </summary>
            /// <param name="parent">TokenExtensionConfig containing necessary context & methods</param>
            public Converters(TokenExtensionConfig parent)
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
