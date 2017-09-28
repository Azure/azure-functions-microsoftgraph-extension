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
    internal static class ExcelClient
    {
        public static async Task<WorkbookTable> GetTableWorkbookAsync(this IGraphServiceClient client, string path, string worksheetName, string tableName)
        {
            return await client
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

        public static async Task<WorkbookRange> GetTableWorkbookRangeAsync(this IGraphServiceClient client, string path, string worksheetName, string tableName)
        {
            return await client
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

        public static async Task<WorkbookRange> GetWorksheetWorkbookAsync(this IGraphServiceClient client, string path, string worksheetName)
        {
            return await client
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

        public static async Task<WorkbookRange> GetWorkSheetWorkbookInRangeAsync(this IGraphServiceClient client, string path, string worksheetName, string range)
        {
            return await client
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

        public static async Task<string[]> GetTableHeaderRowAsync(this IGraphServiceClient client, string path, string tableName)
        {
            var headerRowRange = await client
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

        public static async Task<WorkbookTableRow> PostTableRowAsync(this IGraphServiceClient client, string path, string tableName, JToken row)
        {
            return await client
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

        public static async Task<WorkbookRange> PatchWorksheetAsync(this IGraphServiceClient client, string path, string worksheetName, string range, WorkbookRange newWorkbook)
        {
            return await client
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
