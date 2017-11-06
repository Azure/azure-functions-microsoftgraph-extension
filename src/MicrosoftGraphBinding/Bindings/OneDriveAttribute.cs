// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs
{
    using System.IO;
    using Microsoft.Azure.WebJobs.Description;

    /// <summary>
    /// Binding to O365 OneDrive.
    /// </summary>
    [Binding]
    public class OneDriveAttribute : GraphTokenAttribute
    {
        /// <summary>
        /// Gets or sets path FROM ONEDRIVE ROOT to file
        /// e.g. "Documents/testing.docx"
        /// </summary>
        [AutoResolve]
        public string Path { get; set; }

        public FileAccess? Access { get; set; }

        public OneDriveAttribute()
        {

        }

        public OneDriveAttribute(FileAccess access)
        {
            Access = access;
        }
    }
}