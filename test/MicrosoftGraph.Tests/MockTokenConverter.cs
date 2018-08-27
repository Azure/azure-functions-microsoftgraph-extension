// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.IdentityModel.Tokens;

    internal class MockTokenConverter : IAsyncConverter<TokenAttribute, string>
    {
        private readonly string _secretKey = "dummy_secret_key";
        private readonly JwtSecurityTokenHandler _tokenHandler = new JwtSecurityTokenHandler();

        public Task<string> ConvertAsync(TokenAttribute input, CancellationToken cancellationToken)
        {
            ClaimsIdentity identity = new ClaimsIdentity();
            identity.AddClaim(new Claim(ClaimTypes.Name, "Sample"));
            identity.AddClaim(new Claim("idp", "aad"));
            identity.AddClaim(new Claim("oid", Guid.NewGuid().ToString()));
            identity.AddClaim(new Claim("scp", "read"));

            var descr = new SecurityTokenDescriptor
            {
                Audience = "https://sample.com",
                Issuer = "https://microsoft.graph.com",
                Subject = identity,
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_secretKey)), SecurityAlgorithms.HmacSha256),
                Expires = DateTime.UtcNow.AddHours(1)
            };
            string accessToken = _tokenHandler.CreateJwtSecurityToken(descr).RawData;
            return Task.FromResult(accessToken);
        }
    }
}
