// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Helper for calling onto Excel (MS) Graph
    /// </summary>
    internal class ExcelClient : IExcelClient
    {
        private Task<IGraphServiceClient> _client;

        public ExcelClient(Task<IGraphServiceClient> client)
        {
            _client = client;
        }

        public async Task<WorkbookTable> GetTableWorkbookAsync(string path, string worksheetName, string tableName)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName]
                .Tables[tableName]
                .Request()
                .GetAsync();
        }

        public async Task<WorkbookRange> GetTableWorkbookRangeAsync(string path, string worksheetName, string tableName)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName]
                .Tables[tableName]
                .Range()
                .Request()
                .GetAsync();
        }

        public async Task<WorkbookRange> GetWorksheetWorkbookAsync(string path, string worksheetName)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName]
                .UsedRange()
                .Request()
                .GetAsync();
        }
        public async Task<WorkbookRange> GetWorkSheetWorkbookInRangeAsync(string path, string worksheetName, string range)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName]
                .Range(range)
                .Request()
                .GetAsync();
        }

        public async Task<string[]> GetTableHeaderRowAsync(string path, string tableName)
        {
            var headerRowRange = await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Tables[tableName]
                .HeaderRowRange()
                .Request()
                .GetAsync();
            return headerRowRange.Values.ToObject<string[]>();
        }

        public async Task<WorkbookTableRow> PostTableRowAsync(string path, string tableName, JToken row)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Tables[tableName]
                .Rows
                .Add(null, row)
                .Request()
                .PostAsync();
        }

        public async Task<WorkbookRange> PatchWorksheetAsync(string path, string worksheetName, string range, WorkbookRange newWorkbook)
        {
            return await (await _client)
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName]
                .Range(range)
                .Request()
                .PatchAsync(newWorkbook);
        }
    }
}
