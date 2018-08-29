// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services
{
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;

    internal class ExcelService
    {
        private GraphServiceClientManager _clientProvider;

        public ExcelService(GraphServiceClientManager clientProvider)
        {
            _clientProvider = clientProvider;
        }

        internal async Task<WorkbookTable> GetExcelTableAsync(ExcelAttribute attr, CancellationToken token)
        {
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            return await client.GetTableWorkbookAsync(attr.Path, attr.TableName, token);
        }

        /// <summary>
        /// Returns either an Excel table or entire worksheet depending on user settings
        /// </summary>
        /// <param name="client">GraphServiceClient that makes request</param>
        /// <param name="attr">Contains metadata (path, tablename, worksheet name) </param>
        /// <returns>string [][] containing table contents</returns>
        internal async Task<string[][]> GetExcelRangeAsync(ExcelAttribute attr, CancellationToken token)
        {
            WorkbookRange range;
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            // If TableName is set, then retrieve the contents of a table
            if (attr.TableName != null)
            {
                range = await client.GetTableWorkbookRangeAsync(attr.Path, attr.TableName, token);
            }
            else
            {
                // If TableName is NOT set, then retrieve either the contents or the formulas of the worksheet
                range = await client.GetWorksheetWorkbookAsync(attr.Path, attr.WorksheetName, token);
            }
            return range.Values.ToObject<string[][]>();
        }

        /// <summary>
        /// Returns either an Excel table or entire worksheet depending on user settings
        /// </summary>
        /// <param name="client">GraphServiceClient that makes request</param>
        /// <param name="attr">Contains metadata (path, tablename, worksheet name) </param>
        /// <returns>POCO Array of worksheet or table data</returns>
        internal async Task<T[]> GetExcelRangePOCOAsync<T>(ExcelAttribute attr, CancellationToken token)
        {
            // If TableName is set, then retrieve the contents of a table
            string[][] output = await GetExcelRangeAsync(attr, token);
            string[] header = output[0];
            Dictionary<string, int> dict = new Dictionary<string, int>(); // Map string header value to its index

            // Initialize dictionary
            foreach (var heading in header.Select((value, index) => new { index, value }))
            {
                dict.Add(heading.value, heading.index);
            }

            T[] POCOArray = new T[output.GetLength(0) - 1]; // Init POCO Array size to size of output - header

            // Create array of POCOs from output array; skip header
            foreach (var row in output.Skip(1).Select((value, index) => new { index, value }))
            {
                var POCORow = Activator.CreateInstance(typeof(T), new object[] { });
                var fields = typeof(T).GetProperties();  // Retrieve all of T's fields
                foreach (var field in fields)
                {
                    int indexOfFieldValue;

                    // For each field, find the corresponding index in the output array
                    try
                    {
                        indexOfFieldValue = dict[field.Name];

                        // Then set POCORow's field to the value at that index
                        field.SetValue(POCORow, row.value[indexOfFieldValue]);
                    }
                    catch (KeyNotFoundException)
                    {
                        // If key isn't found in dictionary corresponding to Table's column names, then let user know
                        throw new KeyNotFoundException($"POCO type [{typeof(T)}] contains field [{field.Name}] that was not found in header of Excel table [{attr.TableName}]");
                    }
                }

                POCOArray[row.index] = (T)POCORow;
            }

            return POCOArray;
        }

        /// <summary>
        /// Add row from a Function's dynamic input
        /// </summary>
        /// <param name="client">MS Graph client used to send request</param>
        /// <param name="attr">Excel Attribute with necessary data (workbook name, table name) to build request</param>
        /// <param name="jsonContent">JObject with the data to be added to the table</param>
        /// <returns>WorkbookTableRow that was just added</returns>
        internal async Task AddRowAsync(ExcelAttribute attr, JObject jsonContent, CancellationToken token)
        {
            /*
             * Two options:
             * 1. JObject created from POCO representing strongly typed table -- indicated by "Microsoft.O365Bindings.POCO" being set
             * 2. JObject "values" set to object[][], so simply post an update to specified table -- indicated by "Microsoft.O365Bindings.values" being set
            */
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);

            JToken newRow;
            if (jsonContent[O365Constants.POCOKey] != null)
            {
                string[] headerRow = await client.GetTableHeaderRowAsync(attr.Path, attr.TableName, token);
                jsonContent.Remove(O365Constants.POCOKey); // Remove now unnecessary flag
                newRow = JArray.FromObject(POCOToStringArray(jsonContent, headerRow));
            }
            else if (jsonContent[O365Constants.ValuesKey] != null)
            {
                newRow = jsonContent[O365Constants.ValuesKey];
            }
            else
            {
                throw new KeyNotFoundException($"When appending a row, the '{O365Constants.ValuesKey}' or '{O365Constants.POCOKey}' key must be set");
            }

            await client.PostTableRowAsync(attr.Path, attr.TableName, newRow, token);
        }

        /// <summary>
        /// Update an existing Excel worksheet
        /// Starting at first row, insert the given rows
        /// Overwrites existing data
        /// </summary>
        /// <param name="attr">ExcelAttribute with workbook & worksheet names, starting row & column</param>
        /// <param name="jsonContent">Values with which to update worksheet plus metadata</param>
        /// <returns>WorkbookRange containing updated worksheet</returns>
        internal async Task<WorkbookRange> UpdateWorksheetAsync(ExcelAttribute attr, JObject jsonContent, CancellationToken token)
        {
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            // Retrieve current range of worksheet
            var currentRange = await client.GetWorksheetWorkbookAsync(attr.Path, attr.WorksheetName, token);
            var rowsToBeChanged = int.Parse(jsonContent[O365Constants.RowsKey].ToString());
            var fromTable = !string.IsNullOrEmpty(attr.TableName);
            string newRange = FindNewRange(currentRange.Address, rowsToBeChanged, fromTable);

            // Retrieve old workbook
            WorkbookRange workbook = await client.GetWorkSheetWorkbookInRangeAsync(attr.Path, attr.WorksheetName, newRange, token);

            JToken newRowArray;
            if (jsonContent[O365Constants.POCOKey] != null)
            {
                string[] header = await client.GetTableHeaderRowAsync(attr.Path, attr.TableName, token);
                jsonContent.Remove(O365Constants.POCOKey); // Remove now unnecessary flag
                var newRows = POCOToStringArray(jsonContent, header);
                newRowArray = JArray.FromObject(newRows);
            }
            else
            {
                newRowArray = jsonContent[O365Constants.ValuesKey];
            }

            // Update necessary fields
            PopulateWorkbookWithNewValue(workbook, newRowArray);
            return await client.PatchWorksheetAsync(attr.Path, attr.WorksheetName, newRange, workbook, token);
        }

        /// <summary>
        /// Update a specified column with a specified value
        /// </summary>
        /// <param name="client">GraphServiceClient used to make calls</param>
        /// <param name="attr">ExcelAttribute with path, table name, and worksheet name</param>
        /// <param name="job">JObject with two keys: 'column' and 'value'</param>
        /// <returns>Workbook range containing updated column</returns>
        internal async Task<WorkbookRange> UpdateColumnAsync(ExcelAttribute attr, JObject job, CancellationToken token)
        {
            IGraphServiceClient client = await _clientProvider.GetMSGraphClientFromTokenAttributeAsync(attr, token);
            // The table API only allows updating or adding one row at a time.
            // Instead we update the worksheet range corresponding to the table

            string currentTableRange = (await client.GetTableWorkbookRangeAsync(attr.Path, attr.TableName, token)).Address;

            // Retrieve current worksheet rows
            WorkbookRange currentTableWorkbook = await client.GetWorkSheetWorkbookInRangeAsync(attr.Path, attr.WorksheetName, currentTableRange, token);

            // Update specified column with specified value
            IEnumerable<string[]> values = currentTableWorkbook.Values.ToObject<IEnumerable<string[]>>();
            var enumer = values.GetEnumerator();
            enumer.MoveNext();

            // Try converting column key to digit
            if (int.TryParse(job["column"].ToString(), out int column))
            {
                column = int.Parse(job["column"].ToString());
            }
            else
            {
                // If it's a column heading, need to find corresponding column index
                var headers = enumer.Current;
                column = Array.FindIndex(headers, x => x == job["column"].ToString());
            }

            while (enumer.MoveNext())
            {
                // Update column of this row to specified value
                enumer.Current[column] = job["value"].ToString();
            }

            var updateValues = JArray.FromObject(values);

            PopulateWorkbookWithNewValue(currentTableWorkbook, updateValues);
            return await client.PatchWorksheetAsync(attr.Path, attr.WorksheetName, currentTableRange, currentTableWorkbook, token);
        }

        /// <summary>
        /// Conversion from object[][] to JArray, then set "values" of JObject
        /// </summary>
        /// <param name="rowsArray">2D object array; each row will later be inserted into the Excel table</param>
        /// <returns>JObject with ("values", converted object[][]) pair</returns>
        internal static string CreateRows(object[][] rowsArray)
        {
            // Convert object[]][] to JArray
            JArray rowData = JArray.FromObject(rowsArray);

            // Set "values" key of new JObject
            JObject jsonContent = new JObject();
            jsonContent[O365Constants.ValuesKey] = rowData;

            // Set "rows", "columns" needed if updating entire worksheet
            jsonContent[O365Constants.RowsKey] = rowsArray.GetLength(0);

            try
            {
                // No exception -- array is rectangular by default
                jsonContent[O365Constants.ColsKey] = rowsArray.GetLength(1);
            }
            catch
            {
                // Jagged array -- have to check if the data is rectangular
                int cols = rowsArray[0].Length;
                foreach (object[] row in rowsArray)
                {
                    if (row.GetLength(0) != cols)
                    {
                        throw new DataMisalignedException("The data inserted must be rectangular");
                    }
                }

                jsonContent[O365Constants.ColsKey] = rowsArray[0].Length;
            }

            return jsonContent.ToString();
        }

        /// <summary>
        /// Convert either a single POCO or POCO lists/arrays to a proper string[][] for uploading
        /// </summary>
        /// <param name="jsonContent">jsonContent with the data & metadata</param>
        /// <param name="header">Table header to get order correct</param>
        /// <returns>string[,] that MS Graph will accept</returns>
        private static string[,] POCOToStringArray(JObject jsonContent, string[] header)
        {
            string[,] newRows = null;
            Dictionary<string, int> fields = ConvertToDictionary(header); // Map header value to column index

            if (jsonContent[O365Constants.ValuesKey] != null)
            {
                // Appending multiple rows T[]
                var rows = jsonContent[O365Constants.ValuesKey];
                newRows = new string[rows.Count<JToken>(), fields.Count]; // rows to be appended x columns
                var x = rows.Children<JObject>();

                foreach (var row in rows.Children<JObject>().Select((value, i) => new { i, value }))
                {
                    foreach (KeyValuePair<string, JToken> pair in row.value)
                    {
                        if (fields.ContainsKey(pair.Key))
                        {
                            int index = fields[pair.Key];
                            newRows[row.i, index] = pair.Value.ToString();
                        }
                    }
                }
            }
            else
            {
                // Appending a single row (T)
                newRows = new string[1, fields.Count]; // SDK expects 2D array

                // Initialize column indices of single row with proper value
                foreach (var pair in jsonContent)
                {
                    int index = fields[pair.Key];
                    newRows[0, index] = pair.Value.ToString();
                }
            }

            return newRows;
        }

        /// <summary>
        /// Given a JToken, map each of its keys to their index
        /// </summary>
        /// <param name="jt">JToken whose keys are the headings of a table</param>
        /// <returns>Dictionary mapping heading names to their column index</returns>
        private static Dictionary<string, int> ConvertToDictionary(string[] headerRow)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();

            for(int index = 0; index < headerRow.Length; index++)
            {
                dict[headerRow[index]] = index;
            }

            return dict;
        }

        /// <summary>
        /// Given current range and number of rows to be inserted, determine the new range
        /// </summary>
        /// <param name="range">Current range of table</param>
        /// <param name="rowsToInsert">Number of rows that will be inserted</param>
        /// <param name="ignoreHeader">If true, add one to starting row so as not to overwrite header</param>
        /// <returns>Range that worksheet/table will cover</returns>
        private static string FindNewRange(string range, int rowsToInsert, bool ignoreHeader)
        {
            // Addresses are always in the form "wksht!A3:D4"
            string[] half = range.Split('!'); // separate worksheet name from range
            string worksheetName = half[0];

            string[] rangeArray = half[1].Split(':');  // separate starting row, column from ending row, column
            string startingColumn = new string(rangeArray[0].Where(char.IsLetter).ToArray());
            string startingRow = new string(rangeArray[0].Where(char.IsDigit).ToArray());

            int startingRowInt = int.Parse(startingRow);
            if (ignoreHeader)
            {
                startingRowInt++;
            }

            string endingColumn = new string(rangeArray[1].Where(char.IsLetter).ToArray());
            string endingRow = (startingRowInt + rowsToInsert - 1).ToString();

            string newRange = worksheetName + "!" +
                startingColumn + startingRowInt + ":" +
                endingColumn + endingRow;

            return newRange;
        }

        private static void PopulateWorkbookWithNewValue(WorkbookRange workbook, JToken newValue)
        {
            workbook.Values = newValue;
            workbook.Text = newValue;
            workbook.Formulas = newValue;
            workbook.FormulasLocal = newValue;
            workbook.FormulasR1C1 = newValue;
        }
    }
}
