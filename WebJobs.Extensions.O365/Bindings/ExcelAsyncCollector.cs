// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Collector class used to accumulate and then dispatch requests to MS Graph related to Excel
    /// </summary>
    internal class ExcelAsyncCollector : IAsyncCollector<JObject>
    {
        private readonly ExcelService _manager;
        private readonly ExcelAttribute _attribute;
        private readonly Collection<JObject> _rows;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelAsyncCollector"/> class.
        /// </summary>
        /// <param name="client">GraphServiceClient used to make calls to MS Graph</param>
        /// <param name="attribute">ExcelAttribute containing necessary info about workbook, etc.</param>
        public ExcelAsyncCollector(ExcelService manager, ExcelAttribute attribute)
        {
            _manager = manager;
            _attribute = attribute;
            _rows = new Collection<JObject>();
        }

        /// <summary>
        /// Add a JObject to the list of JObjects that need to be processed
        /// </summary>
        /// <param name="item">JObject to be added (contains table, worksheet, etc. data)</param>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>Task representing the addition of the JObject</returns>
        public Task AddAsync(JObject item, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (item == null)
            {
                throw new ArgumentNullException("No row item");
            }

            this._rows.Add(item);
            return Task.CompletedTask;
        }

        /// <summary>
        /// Execute all of the JObjects in the collector
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
                if(_rows.FirstOrDefault(row => row["column"] != null && row["value"] != null) != null)
                {
                    foreach (var row in this._rows)
                    {
                        row[O365Constants.RowsKey] = 1;
                        row[O365Constants.ColsKey] = row.Children().Count();
                        if (row["column"] != null && row["value"] != null)
                        {
                            await _manager.UpdateColumn(this._attribute, row);
                        }
                        else
                        {
                            // Update whole worksheet
                            await _manager.UpdateWorksheet(this._attribute, row);
                        }
                    }
                } else if (_rows.Count > 0)
                {
                    // Update whole worksheet at once
                    JObject consolidatedRows = GetConsolidatedRows(_rows);
                    await _manager.UpdateWorksheet(_attribute, consolidatedRows);
                }
            }
            else
            {
                // DEFAULT: Append (rows to specific table)
                foreach (var row in this._rows)
                {
                    await _manager.AddRow(this._attribute, row);
                }
            }

            this._rows.Clear();
        }

        public JObject GetConsolidatedRows(Collection<JObject> rows)
        {
            JObject consolidatedRows = new JObject();
            if (rows.Count > 0)
            {
                // List<T> -> JArray
                consolidatedRows[O365Constants.ValuesKey] = JArray.FromObject(rows);

                // Set rows, columns needed if updating entire worksheet
                consolidatedRows[O365Constants.RowsKey] = rows.Count;

                // No exception -- array is rectangular by default
                consolidatedRows[O365Constants.ColsKey] = rows[0].Children().Count();

                // Set POCO key to indicate that the values need to be ordered to match the header of the existing table
                consolidatedRows[O365Constants.POCOKey] = rows[0][O365Constants.POCOKey];
            }
            return consolidatedRows;
        }
    }
}
