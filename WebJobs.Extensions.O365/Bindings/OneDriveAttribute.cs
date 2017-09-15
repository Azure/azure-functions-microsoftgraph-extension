﻿// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
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
        public string Path { get; set; }
    }
}