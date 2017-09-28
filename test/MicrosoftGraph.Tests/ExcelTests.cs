// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Graph;
    using Moq;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Xunit;

    public class ExcelTests
    {
        private static WorkbookTable finalTable;
        private static string[][] finalRange;
        private static SamplePoco[] finalRangePocoArray;
        private static List<SamplePoco> finalRangePocoList;

        private const string path = "sample/path";
        private const string tableName = "tableName";
        private const string worksheetName = "worksheetName";
        private const string fullOldTableAddress = "worksheetName!A1:B5";
        private const string newTableAddressMinusHeader = "worksheetName!A2:B3";

        [Fact]
        public static async Task Input_WorkbookTableObject_ReturnsExpectedValue()
        {
            WorkbookTable table = new WorkbookTable();
            var clientMock = new Mock<IGraphServiceClient>();
            clientMock.MockGetTableWorkbookAsync(table);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>(clientMock, "ExcelInputFunctions.GetWorkbookTable");

            var expectedResult = table;
            Assert.Equal(expectedResult, finalTable);
            ResetState();
        }

        [Fact]
        public static async Task Input_WorkbookTableAsJaggedStringArray_ReturnsExpectedValue()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            string[][] range = GetRangeAsJaggedStringArray();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(range)
            };
            clientMock.MockGetTableWorkbookRangeAsync(workbookRange);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>(clientMock, "ExcelInputFunctions.GetExcelTableRange");

            var expectedResult = range;
            Assert.Equal(expectedResult, finalRange);
            ResetState();
        }

        [Fact]
        public static async Task Input_WorksheetRangeAsJaggedStringArray_ReturnsExpectedValue()
        {
            var range = GetRangeAsJaggedStringArray();
            var clientMock = GetWorksheetClientMock(range);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>(clientMock, "ExcelInputFunctions.GetExcelWorksheetRange");

            var expectedResult = range;
            Assert.Equal(expectedResult, finalRange);
            ResetState();
        }

        [Fact]
        public static async Task Input_WorksheetRangeAsPocoArray_ReturnsExpectedValue()
        {
            var range = GetRangeAsJaggedStringArray();
            var clientMock = GetWorksheetClientMock(range);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>(clientMock, "ExcelInputFunctions.GetExcelWorksheetRangePocoArray");

            var expectedResult = JsonConvert.SerializeObject(GetRangeAsPocoArray());
            Assert.Equal(expectedResult, JsonConvert.SerializeObject(finalRangePocoArray));
            ResetState();
        }

        [Fact]
        public static async Task Input_WorksheetRangeAsPocoList_ReturnsExpectedValue()
        {
            string[][] range = GetRangeAsJaggedStringArray();
            var clientMock = GetWorksheetClientMock(range);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>(clientMock, "ExcelInputFunctions.GetExcelWorksheetRangePocoList");

            var expectedResult = JsonConvert.SerializeObject(GetRangeAsPocoList());
            Assert.Equal(expectedResult, JsonConvert.SerializeObject(finalRangePocoList));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowsWithStringJaggedArray_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>(clientMock, "ExcelOutputFunctions.AppendRowJaggedArray");

            clientMock.VerifyPostTableRowAsync(path, tableName, token => JTokenEqualsAllButHeaderRow(token));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowWithPoco_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>(clientMock, "ExcelOutputFunctions.AppendRowPoco");

            clientMock.VerifyPostTableRowAsync(path, tableName, token => JTokenEqualsAllButHeaderRow(token));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowsWithPocoList_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>(clientMock, "ExcelOutputFunctions.AppendRowPocoList");

            clientMock.VerifyPostTableRowAsync(path, tableName, token => JTokenEqualsAllButHeaderRow(token));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowsWithPocoArray_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>(clientMock, "ExcelOutputFunctions.AppendRowPocoArray");

            clientMock.VerifyPostTableRowAsync(path, tableName, token => JTokenEqualsAllButHeaderRow(token));
            ResetState();
        }

        [Fact]
        public static async Task Update_WorksheetWithTable_SendsPatchWithProperValues()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            var headerRow = new WorkbookRange()
            {
                Values = JToken.FromObject(GetHeaderRow())
            };
            var address = new WorkbookRange()
            {
                Address = fullOldTableAddress,
            };
            var workbookInRange = new WorkbookRange();
            clientMock.MockGetWorksheetWorkbookAsync(address);
            clientMock.MockGetTableHeaderRowAsync(headerRow);
            clientMock.MockGetWorkSheetWorkbookInRangeAsync(workbookInRange);
            clientMock.MockPatchWorksheetAsync(null);

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>(clientMock, "ExcelOutputFunctions.UpdateWorksheet");
            var samplePocos = GetRangeAsPocoArray();
            clientMock.VerifyPatchWorksheetAsync(path, worksheetName, newTableAddressMinusHeader, range => JTokenEqualsPocos(range.Values, samplePocos));
            ResetState();
        }

        private static Mock<IGraphServiceClient> AppendClientMock()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            string[] headerRow = GetHeaderRow();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(headerRow)
            };
            clientMock.MockGetTableHeaderRowAsync(workbookRange);
            clientMock.MockPostTableRowAsyc(null);
            return clientMock;
        }

        private static Mock<IGraphServiceClient> GetWorksheetClientMock(object range)
        {
            var clientMock = new Mock<IGraphServiceClient>();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(range)
            };
            clientMock.MockGetWorksheetWorkbookAsync(workbookRange);
            return clientMock;
        }

        private static string[] GetHeaderRow()
        {
            return GetRangeAsJaggedStringArray()[0];
        }

        private static string[][] GetAllButHeaderRow()
        {
            var range = GetRangeAsJaggedStringArray();
            return new string[][] { range[1], range[2] };
        }

        private static object[][] ConvertJaggedArrayType(string[][] jaggedStringArray)
        {
            object[][] jaggedObjectArray = new object[jaggedStringArray.Length][];
            for(var rowIndex = 0; rowIndex < jaggedStringArray.Length; rowIndex++)
            {
                jaggedObjectArray[rowIndex] = new object[jaggedStringArray[rowIndex].Length];
                for(var columnIndex = 0; columnIndex < jaggedObjectArray[rowIndex].Length; columnIndex++)
                {
                    jaggedObjectArray[rowIndex][columnIndex] = jaggedStringArray[rowIndex][columnIndex];
                }
            }
            return jaggedObjectArray;
        }

        private static bool JTokenEqualsAllButHeaderRow(JToken value)
        {
            var allButHeaderRow = ConvertJaggedArrayType(GetAllButHeaderRow());
            var values = value.ToObject<JObject>()[O365Constants.ValuesKey].ToObject<object[][]>();
            for(int rowIndex = 0; rowIndex < allButHeaderRow.Length; rowIndex++)
            {
                for(int columnIndex = 0; columnIndex < allButHeaderRow[rowIndex].Length; columnIndex++)
                {
                    if(allButHeaderRow[rowIndex][columnIndex] != values[rowIndex][columnIndex])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private static bool JTokenEqualsPocos(JToken token, IEnumerable<SamplePoco> pocos)
        {
            string[,] stringMultiArray = token.ToObject<JArray>().ToObject<string[,]>();
            int index = 0;
            foreach (var poco in pocos)
            {
                if(!string.Equals(stringMultiArray[index, 0], poco.Name) || !string.Equals(stringMultiArray[index,1], poco.Value))
                {
                    return false;
                }
                index++;
            }
            return true;
        }

        private static string[][] GetRangeAsJaggedStringArray()
        {
            string[][] range = new string[3][];
            range[0] = new string[] { "Name", "Value" };
            range[1] = new string[] { "Name1", "Value1" };
            range[2] = new string[] { "Name2", "Value2" };
            return range;
        }

        private static SamplePoco[] GetRangeAsPocoArray()
        {
            return new SamplePoco[]
            {
                new SamplePoco() {Name = "Name1", Value = "Value1"},
                new SamplePoco() {Name = "Name2", Value = "Value2"}
            };
        }

        private static List<SamplePoco> GetRangeAsPocoList()
        {
            return new List<SamplePoco>(GetRangeAsPocoArray());
        }

        private static void ResetState()
        {
            finalTable = null;
            finalRange = null;
            finalRangePocoArray = null;
            finalRangePocoList = null;
        }

        private class ExcelInputFunctions
        {
            public void GetWorkbookTable(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] WorkbookTable table)
            {
                finalTable = table;
            }

            public void GetExcelTableRange(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] string[][] range)
            {
                finalRange = range;
            }

            public void GetExcelWorksheetRange(
                [Excel(
                Path = path,
                WorksheetName = worksheetName)] string[][] range)
            {
                finalRange = range;
            }

            public void GetExcelWorksheetRangePocoArray(
                [Excel(
                Path = path,
                WorksheetName = worksheetName)] SamplePoco[] range)
            {
                finalRangePocoArray = range;
            }

            public void GetExcelWorksheetRangePocoList(
                [Excel(
                Path = path,
                WorksheetName = worksheetName)] List<SamplePoco> range)
            {
                finalRangePocoList = range;
            }
        }

        private class ExcelOutputFunctions 
        {
            public void AppendRowJaggedArray(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out object[][] range)
            {
                range = ConvertJaggedArrayType(GetAllButHeaderRow());
            }

            public void AppendRowPoco(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out SamplePoco row)
            {
                row = GetRangeAsPocoArray()[0];
            }

            public void AppendRowPocoList(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out List<SamplePoco> rows)
            {
                rows = GetRangeAsPocoList();
            }

            public void AppendRowPocoArray(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out SamplePoco[] rows)
            {
                rows = GetRangeAsPocoArray();
            }

            public void UpdateWorksheet([Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName,
                UpdateType = "Update")] ICollector<SamplePoco> rows)
            {
                foreach (var row in GetRangeAsPocoList())
                {
                    rows.Add(row);
                }
            }
        }

        private class SamplePoco
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }
    }
}
