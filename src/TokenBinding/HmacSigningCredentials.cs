// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.AuthTokens
{
    using System;
    using Microsoft.IdentityModel.Tokens;

    internal class HmacSigningCredentials : SigningCredentials
    {
        public HmacSigningCredentials(string base64EncodedKey)
            : this(ParseKeyString(base64EncodedKey))
        { }

        public HmacSigningCredentials(byte[] key)
            : base(new SymmetricSecurityKey(key), CreateSignatureAlgorithm(key))
        {
        }

        protected static byte[] ParseKeyString(string keyString)
        {
            if (string.IsNullOrEmpty(keyString))
            {
                return new byte[0];
            }
            else if (IsHexString(keyString))
            {
                return HexStringToByteArray(keyString);
            }
            else
            {
                return Convert.FromBase64String(keyString);
            }
        }

        protected static bool IsHexString(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                throw new ArgumentNullException(nameof(input));
            }

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                bool isHexChar = (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
                if (!isHexChar)
                {
                    return false;
                }
            }

            return true;
        }

        protected static byte[] HexStringToByteArray(string hexString)
        {
            byte[] bytes = new byte[hexString.Length / 2];
            for (int i = 0; i < hexString.Length; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }

            return bytes;
        }

        protected static string CreateSignatureAlgorithm(byte[] key)
        {
            if (key.Length <= 32)
            {
                return Algorithms.HmacSha256Signature;
            }
            else if (key.Length <= 48)
            {
                return Algorithms.HmacSha384Signature;
            }
            else
            {
                return Algorithms.HmacSha512Signature;
            }
        }

        /// <summary>
        /// Key value pairs (algorithm name, w3.org link)
        /// </summary>
        protected static class Algorithms
        {
            public const string HmacSha256Signature = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha256";
            public const string HmacSha384Signature = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha384";
            public const string HmacSha512Signature = "http://www.w3.org/2001/04/xmldsig-more#hmac-sha512";
        }
    }
}
