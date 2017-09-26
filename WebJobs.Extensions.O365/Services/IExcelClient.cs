// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests")]
namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    interface IExcelClient
    {
        Task<WorkbookTable> GetTableWorkbookAsync(string path, string worksheetName, string tableName);

        Task<WorkbookRange> GetTableWorkbookRangeAsync(string path, string worksheetName, string tableName);

        Task<WorkbookRange> GetWorksheetWorkbookAsync(string path, string worksheetName);

        Task<WorkbookRange> GetWorkSheetWorkbookInRangeAsync(string path, string worksheetName, string range);

        Task<string[]> GetTableHeaderRowAsync(string path, string tableName);

        Task<WorkbookTableRow> PostTableRowAsync(string path, string tableName, JToken row);

        Task<WorkbookRange> PatchWorksheetAsync(string path, string worksheetName, string range, WorkbookRange newWorkbook);
    }
}
