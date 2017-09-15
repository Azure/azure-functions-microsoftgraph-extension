// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.Bindings
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Description;

    /// <summary>
    /// Attribute used to describe Excel files
    /// Files must be on OneDrive
    /// </summary>
    [Binding]
    public class ExcelAttribute : GraphTokenAttribute
    {
        /// <summary>
        /// Gets or sets path FROM ONEDRIVE ROOT to Excel file
        /// e.g. "Documents/TestSheet.xlsx"
        /// </summary>
        [AutoResolve]
        public string Path { get; set; }

        /// <summary>
        /// Gets or sets name of Excel table
        /// </summary>
        [AutoResolve]
        public string TableName { get; set; }

        /// <summary>
        /// Gets or sets worksheet name
        /// </summary>
        [AutoResolve]
        public string WorksheetName { get; set; }

        /// <summary>
        /// Gets or sets used when reading/writing from/to file
        /// Append (row) or Update (whole worksheet, specific column)
        /// </summary>
        [AutoResolve]
        public string UpdateType { get; set; }
    }
}