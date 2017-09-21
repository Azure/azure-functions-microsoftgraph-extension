// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph
{
    using System;
    using System.Collections.ObjectModel;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Collector class used to accumulate and then dispatch requests to MS Graph related to Excel
    /// </summary>
    internal class ExcelAsyncCollector : IAsyncCollector<JObject>
    {
        private readonly GraphServiceClient client;
        private readonly ExcelAttribute attribute;
        private readonly Collection<JObject> rows = new Collection<JObject>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelAsyncCollector"/> class.
        /// </summary>
        /// <param name="client">GraphServiceClient used to make calls to MS Graph</param>
        /// <param name="attribute">ExcelAttribute containing necessary info about workbook, etc.</param>
        public ExcelAsyncCollector(GraphServiceClient client, ExcelAttribute attribute)
        {
            this.client = client;
            this.attribute = attribute;
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

            this.rows.Add(item);
            return Task.CompletedTask;
        }

        /// <summary>
        /// Execute all of the JObjects in the collector
        /// </summary>
        /// <param name="cancellationToken">Used to propagate notifications</param>
        /// <returns>Task representing the flushing of the collector</returns>
        public async Task FlushAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            // Distinguish between appending and updating
            if (this.attribute.UpdateType != null && this.attribute.UpdateType == "Update")
            {
                // Update whole worksheet, whole table, or specific column

                // Updating specific column requires Column, Value set
                foreach (var row in this.rows)
                {
                    if (row["column"] != null && row["value"] != null)
                    {
                        await this.client.UpdateColumn(this.attribute, row);
                    }
                    else
                    {
                        // Update whole worksheet
                        await this.client.UpdateWorksheet(this.attribute, row);
                    }
                }
            }
            else
            {
                // DEFAULT: Append (rows to specific table)
                foreach (var row in this.rows)
                {
                    await this.client.AddRow(this.attribute, row);
                }
            }

            this.rows.Clear();
        }
    }
}
