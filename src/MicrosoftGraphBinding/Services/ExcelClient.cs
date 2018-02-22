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
        public static async Task<WorkbookTable> GetTableWorkbookAsync(this IGraphServiceClient client, string path, string tableName)
        {
            return await client
                .GetWorkbookTableRequest(path, tableName)
                .Request()
                .GetAsync();
        }

        public static async Task<WorkbookRange> GetTableWorkbookRangeAsync(this IGraphServiceClient client, string path, string tableName)
        {
            return await client
                .GetWorkbookTableRequest(path, tableName)
                .Range()
                .Request()
                .GetAsync();
        }

        public static async Task<WorkbookRange> GetWorksheetWorkbookAsync(this IGraphServiceClient client, string path, string worksheetName)
        {
            return await client
                .GetWorkbookWorksheetRequest(path, worksheetName)
                .UsedRange()
                .Request()
                .GetAsync();
        }

        public static async Task<WorkbookRange> GetWorkSheetWorkbookInRangeAsync(this IGraphServiceClient client, string path, string worksheetName, string range)
        {
            return await client
                .GetWorkbookWorksheetRequest(path, worksheetName)
                .Range(range)
                .Request()
                .GetAsync();
        }

        public static async Task<string[]> GetTableHeaderRowAsync(this IGraphServiceClient client, string path, string tableName)
        {
            var headerRowRange = await client
                .GetWorkbookTableRequest(path, tableName)
                .HeaderRowRange()
                .Request()
                .GetAsync();
            return headerRowRange.Values.ToObject<string[][]>()[0]; //header row array is embedded as the only element in its own array
        }

        public static async Task<WorkbookTableRow> PostTableRowAsync(this IGraphServiceClient client, string path, string tableName, JToken row)
        {
            return await client
                .GetWorkbookTableRequest(path, tableName)
                .Rows
                .Add(null, row)
                .Request()
                .PostAsync();
        }

        public static async Task<WorkbookRange> PatchWorksheetAsync(this IGraphServiceClient client, string path, string worksheetName, string range, WorkbookRange newWorkbook)
        {
            return await client
                .GetWorkbookWorksheetRequest(path, worksheetName)
                .Range(range)
                .Request()
                .PatchAsync(newWorkbook);
        }


        internal static IWorkbookTableRequestBuilder GetWorkbookTableRequest(this IGraphServiceClient client, string path, string tableName)
        {
            return client
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Tables[tableName];
        }

        internal static IWorkbookWorksheetRequestBuilder GetWorkbookWorksheetRequest(this IGraphServiceClient client, string path, string worksheetName)
        {
            return client
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName];
        }
    }
}
