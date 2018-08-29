// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Collector class used to accumulate and then dispatch requests to MS Graph related to Excel
    /// </summary>
    internal class ExcelAsyncCollector : IAsyncCollector<string>
    {
        private readonly ExcelService _service;
        private readonly ExcelAttribute _attribute;
        private readonly List<JObject> _rows;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelAsyncCollector"/> class.
        /// </summary>
        /// <param name="client">GraphServiceClient used to make calls to MS Graph</param>
        /// <param name="attribute">ExcelAttribute containing necessary info about workbook, etc.</param>
        public ExcelAsyncCollector(ExcelService manager, ExcelAttribute attribute)
        {
            _service = manager;
            _attribute = attribute;
            _rows = new List<JObject>();
        }

        /// <summary>
        /// Add a string representation of a JObject to the list of items that need to be processed
        /// </summary>
        /// <param name="item">JSON string to be added (contains table, worksheet, etc. data)</param>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>Task representing the addition of the item</returns>
        public Task AddAsync(string item, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }
            JToken parsedToken = JToken.Parse(item);
            if (parsedToken is JObject)
            {
                _rows.Add(parsedToken as JObject);
            }
            else
            {
                //JavaScript Array, so add metadata manually
                var array = parsedToken as JArray;
                bool arrayIsNested = array.All(element => element is JArray);
                if (arrayIsNested)
                {
                    var consolidatedRow = new JObject();
                    consolidatedRow[O365Constants.ValuesKey] = array;
                    consolidatedRow[O365Constants.RowsKey] = array.Count;
                    // No exception -- array is rectangular by default
                    consolidatedRow[O365Constants.ColsKey] = array[0].Children().Count();
                    _rows.Add(consolidatedRow);
                }
                else
                {
                    throw new InvalidOperationException("Only nested arrays are supported for Excel output bindings.");
                }
                
            }
            
            return Task.CompletedTask;
        }

        /// <summary>
        /// Send all of the items in the collector to Microsoft Graph API
        /// </summary>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>Task representing the flushing of the collector</returns>
        public async Task FlushAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            if (_rows.Count == 0)
            {
                return;
            }
            // Distinguish between appending and updating
            if (this._attribute.UpdateType != null && this._attribute.UpdateType == "Update")
            {
                if (_rows.FirstOrDefault(row => row["column"] != null && row["value"] != null) != null)
                {
                    foreach (var row in this._rows)
                    {
                        row[O365Constants.RowsKey] = 1;
                        row[O365Constants.ColsKey] = row.Children().Count();
                        if (row["column"] != null && row["value"] != null)
                        {
                            await _service.UpdateColumnAsync(this._attribute, row, cancellationToken);
                        }
                        else
                        {
                            // Update whole worksheet
                            await _service.UpdateWorksheetAsync(this._attribute, row, cancellationToken);
                        }
                    }
                }
                else if (_rows.Count > 0)
                {
                    // Update whole worksheet at once
                    JObject consolidatedRows = GetConsolidatedRows(_rows);
                    await _service.UpdateWorksheetAsync(_attribute, consolidatedRows, cancellationToken);
                }
            }
            else
            {
                // DEFAULT: Append (rows to specific table)
                foreach (var row in this._rows)
                {
                    await _service.AddRowAsync(this._attribute, row, cancellationToken);
                }
            }

            this._rows.Clear();
        }

        public JObject GetConsolidatedRows(List<JObject> rows)
        {
            JObject consolidatedRows = new JObject();
            if (rows.Count > 0 && rows[0][O365Constants.ValuesKey] == null)
            {
                //Each row is a POCO, so make it one value, and consolidate into one row
                consolidatedRows[O365Constants.ValuesKey] = JArray.FromObject(rows);
                consolidatedRows[O365Constants.RowsKey] = rows.Count;
                // No exception -- array is rectangular by default
                consolidatedRows[O365Constants.ColsKey] = rows[0].Children().Count();
                // Set POCO key to indicate that the values need to be ordered to match the header of the existing table
                consolidatedRows[O365Constants.POCOKey] = rows[0][O365Constants.POCOKey];
            }
            else if (rows.Count == 1 && rows[0][O365Constants.ValuesKey] != null)
            {
                return rows[0];
            }
            else
            {
                var rowsAsString = $"[ {string.Join(", ", rows.Select(jobj => jobj.ToString()))} ]";
                throw new InvalidOperationException($"Could not consolidate the following rows: {rowsAsString}");
            }
            return consolidatedRows;
        }
    }
}
