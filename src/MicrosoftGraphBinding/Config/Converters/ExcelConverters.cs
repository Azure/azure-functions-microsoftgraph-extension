// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Config.Converters
{
    using System.Linq;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Newtonsoft.Json.Linq;
    using Microsoft.Graph;

    internal class ExcelConverters
    {
        internal class ExcelConverter : 
            IAsyncConverter<ExcelAttribute, string[][]>,
            IAsyncConverter<ExcelAttribute, WorkbookTable>,
            IAsyncConverter<ExcelAttribute, IAsyncCollector<string>>
        {
            private readonly ServiceManager _serviceManager;

            public ExcelConverter(ServiceManager serviceManager)
            {
                _serviceManager = serviceManager;
            }

            async Task<IAsyncCollector<string>> IAsyncConverter<ExcelAttribute, IAsyncCollector<string>>.ConvertAsync(ExcelAttribute attr, CancellationToken token)
            {
                var manager = await _serviceManager.GetExcelService(attr);
                return new ExcelAsyncCollector(manager, attr);
            }

            async Task<string[][]> IAsyncConverter<ExcelAttribute, string[][]>.ConvertAsync(ExcelAttribute attr, CancellationToken cancellationToken)
            {
                var service = await _serviceManager.GetExcelService(attr);
                return await service.GetExcelRangeAsync(attr);
            }

            async Task<WorkbookTable> IAsyncConverter<ExcelAttribute, WorkbookTable>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var service = await _serviceManager.GetExcelService(input);
                return await service.GetExcelTable(input);
            }
        }

        internal class ExcelGenericsConverter<T> : IAsyncConverter<ExcelAttribute, T[]>, 
            IConverter<T, string>
        {
            private readonly ServiceManager _serviceManager;

            /// <summary>
            /// Initializes a new instance of the <see cref="ExcelGenericsConverter{T}"/> class.
            /// </summary>
            /// <param name="parent">O365Extension to which the result of the request for data will be returned</param>
            public ExcelGenericsConverter(ServiceManager serviceManager)
            {
                this._serviceManager = serviceManager;
            }

            async Task<T[]> IAsyncConverter<ExcelAttribute, T[]>.ConvertAsync(ExcelAttribute input, CancellationToken cancellationToken)
            {
                var manager = await _serviceManager.GetExcelService(input);
                return await manager.GetExcelRangePOCOAsync<T>(input);
            }

            /// <summary>
            /// Convert from POCO -> JObject (either row or rows)
            /// </summary>
            /// <param name="input">POCO input from fx</param>
            /// <returns>JObject with proper keys set</returns>
            public string Convert(T input)
            {
                // handle T[]
                if (typeof(T).IsArray)
                {
                    var array = input as object[];
                    return ConvertEnumerable(array).ToString();
                }
                else
                {
                    // handle T
                    JObject data = JObject.FromObject(input);
                    data[O365Constants.POCOKey] = true; // Set Microsoft.O365Bindings.POCO flag to indicate that data is from POCO (vs. object[][])

                    return data.ToString();
                }
            }

            /// <summary>
            /// Convert from List<POCO> -> JObject
            /// </summary>
            /// <param name="input">POCO input from fx</param>
            /// <returns>JObject with proper keys set</returns>
            public JObject Convert(List<T> input)
            {
                return ConvertEnumerable(input);
            }

            private JObject ConvertEnumerable<U>(IEnumerable<U> input)
            {
                JObject jsonContent = new JObject();

                JArray rowData = JArray.FromObject(input);

                // List<T> -> JArray
                jsonContent[O365Constants.ValuesKey] = rowData;

                // Set rows, columns needed if updating entire worksheet
                jsonContent[O365Constants.RowsKey] = rowData.Count();

                // No exception -- array is rectangular by default
                jsonContent[O365Constants.ColsKey] = rowData.First.Count();

                // Set POCO key to indicate that the values need to be ordered to match the header of the existing table
                jsonContent[O365Constants.POCOKey] = true;

                return jsonContent;
            }
        }
    }
}
