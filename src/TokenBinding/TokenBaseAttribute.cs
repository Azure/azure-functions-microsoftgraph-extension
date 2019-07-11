// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using Microsoft.Azure.WebJobs.Description;
    using Microsoft.Azure.WebJobs.Host.Bindings;
    using Microsoft.Azure.WebJobs.Host.Bindings.Path;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Azure.WebJobs.Extensions.AuthTokens;
    using Microsoft.Extensions.Primitives;
    using Microsoft.Net.Http.Headers;

    public abstract class TokenBaseAttribute : Attribute
    {
        [AutoResolve(ResolutionPolicyType = typeof(EasyAuthAccessTokenResolutionPolicy))]
        public string EasyAuthAccessToken { get; set; } = "auto"; // Needs to be set to non-null value to be resolved

        /// <summary>
        /// Gets or sets a resource for a token exchange. Optional
        /// </summary>
        public string AadResource { get; set; }

        /// <summary>
        /// Gets or sets how to determine identity. Required.
        /// </summary>
        public TokenIdentityMode Identity { get; set; }

        internal class EasyAuthAccessTokenResolutionPolicy : IResolutionPolicy
        {
            public string TemplateBind(PropertyInfo propInfo, Attribute resolvedAttribute, BindingTemplate bindingTemplate, IReadOnlyDictionary<string, object> bindingData)
            {
                var tokenAttribute = resolvedAttribute as TokenBaseAttribute;
                if (tokenAttribute == null)
                {
                    throw new InvalidOperationException($"Can not use {nameof(EasyAuthAccessTokenResolutionPolicy)} as a resolution policy for an attribute that does not implement {nameof(TokenBaseAttribute)}");
                }

                if (tokenAttribute.Identity != TokenIdentityMode.UserFromRequest)
                {
                    // No other modes require this field
                    return null;
                }

                if (!(bindingData.ContainsKey("$request") && bindingData["$request"] is HttpRequest))
                {
                    throw new InvalidOperationException($"Can not use {nameof(TokenIdentityMode.UserFromRequest)} mode of {resolvedAttribute.GetType()} with a non-HTTP triggered function.");
                }

                var request = (HttpRequest)bindingData["$request"];
                return GetEasyAuthAccessToken(request);
            }

            private string GetEasyAuthAccessToken(HttpRequest request)
            {
                string errorMessage = "Can not find an access token for the user. Verify that this endpoint is protected by Azure App Service Authentication/Authorization.";
                if (request.Headers.TryGetValue(Constants.EasyAuthAadAccessTokenHeader, out StringValues accessTokenHeaderValue))
                {
                    return accessTokenHeaderValue.ToString();
                }

                if (request.Headers.TryGetValue(HeaderNames.Authorization, out StringValues authorizationHeaderValue))
                {
                    var bearerTokenString = authorizationHeaderValue.ToString();
                    string[] bearerTokenComponents = bearerTokenString.Split(' ');
                    if (bearerTokenComponents.Length != 2 && bearerTokenComponents[0] != "Bearer")
                    {
                        throw new InvalidOperationException(errorMessage);
                    }
                    return bearerTokenComponents[1];
                }

                throw new InvalidOperationException(errorMessage);
            }
        }

    }
}
