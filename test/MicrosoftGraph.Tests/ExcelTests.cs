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

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>("ExcelInputFunctions.GetWorkbookTable", clientMock);

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

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>("ExcelInputFunctions.GetExcelTableRange", clientMock);

            var expectedResult = range;
            Assert.Equal(expectedResult, finalRange);
            ResetState();
        }

        [Fact]
        public static async Task Input_WorksheetRangeAsJaggedStringArray_ReturnsExpectedValue()
        {
            var range = GetRangeAsJaggedStringArray();
            var clientMock = GetWorksheetClientMock(range);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>("ExcelInputFunctions.GetExcelWorksheetRange", clientMock);

            var expectedResult = range;
            Assert.Equal(expectedResult, finalRange);
            ResetState();
        }

        [Fact]
        public static async Task Input_WorksheetRangeAsPocoArray_ReturnsExpectedValue()
        {
            var range = GetRangeAsJaggedStringArray();
            var clientMock = GetWorksheetClientMock(range);

            await CommonUtilities.ExecuteFunction<ExcelInputFunctions>("ExcelInputFunctions.GetExcelWorksheetRangePocoArray", clientMock);

            var expectedResult = JsonConvert.SerializeObject(GetRangeAsPocoArray());
            Assert.Equal(expectedResult, JsonConvert.SerializeObject(finalRangePocoArray));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowsWithStringJaggedArray_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.AppendRowJaggedArray", clientMock);

            string[][] allRows = GetAllButHeaderRow();
            clientMock.VerifyPostTableRowAsync(path, tableName, token => ValuesEqual(allRows, token));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowWithPoco_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.AppendRowPoco", clientMock);
            string[][] allRows = GetAllButHeaderRow();
            clientMock.VerifyPostTableRowAsync(path, tableName, token => ValuesEqual(allRows[0], token));
            ResetState();
        }

        [Fact]
        public static async Task Append_RowsWithPocoCollector_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.AppendRowPocoArray", clientMock);

            string[][] allRows = GetAllButHeaderRow();
            foreach (var row in allRows)
            {
                clientMock.VerifyPostTableRowAsync(path, tableName, token => ValuesEqual(row, token));
            }
            ResetState();
        }

        [Fact]
        public static async Task Append_RowsNode_SendsPostWithProperValues()
        {
            var clientMock = AppendClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.AppendRowNode", clientMock);

            string[][] allRows = GetAllButHeaderRow();
            clientMock.VerifyPostTableRowAsync(path, tableName, token => ValuesEqual(allRows, token));
            ResetState();
        }

        [Fact]
        public static async Task Update_WorksheetWithJaggedArray_SendsPatchWithProperValues()
        {
            var clientMock = UpdateClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.UpdateWorksheetJaggedArray", clientMock);
            string[][] allRows = GetAllButHeaderRow();
            clientMock.VerifyPatchWorksheetAsync(path, worksheetName, newTableAddressMinusHeader, range => JTokenEqualsJaggedArray(range.Values, allRows));

            ResetState();
        }

        [Fact]
        public static async Task Update_WorksheetWithTable_SendsPatchWithProperValues()
        {
            var clientMock = UpdateClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.UpdateWorksheet", clientMock);
            string[][] allRows = GetAllButHeaderRow();
            clientMock.VerifyPatchWorksheetAsync(path, worksheetName, newTableAddressMinusHeader, range => JTokenEqualsJaggedArray(range.Values, allRows));
            ResetState();
        }

        [Fact]
        public static async Task Update_WorksheetNode_SendsPatchWithProperValues()
        {
            var clientMock = UpdateClientMock();

            await CommonUtilities.ExecuteFunction<ExcelOutputFunctions>("ExcelOutputFunctions.UpdateWorksheetNode", clientMock);
            string[][] allRows = GetAllButHeaderRow();
            clientMock.VerifyPatchWorksheetAsync(path, worksheetName, newTableAddressMinusHeader, range => JTokenEqualsJaggedArray(range.Values, allRows));
            ResetState();
        }

        private static Mock<IGraphServiceClient> AppendClientMock()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            string[][] headerRow = new string[][] { GetHeaderRow() };
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(headerRow)
            };
            clientMock.MockGetTableHeaderRowAsync(workbookRange);
            clientMock.MockPostTableRowAsyc(null);
            return clientMock;
        }

        private static Mock<IGraphServiceClient> UpdateClientMock()
        {
            var clientMock = new Mock<IGraphServiceClient>();
            string[][] headerValues = new string[][] { GetHeaderRow() };
            var headerRow = new WorkbookRange()
            {
                Values = JToken.FromObject(headerValues)
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

        private static bool ValuesEqual(string[] expectedValue, JToken actualValue)
        {
            var allButHeaderRow = ConvertJaggedArrayType(GetAllButHeaderRow());
            var value = actualValue.ToObject<JArray>().ToObject<string[][]>()[0];
            for(int i = 0; i < expectedValue.Length; i++)
            {
                if(expectedValue[i] != value[i])
                {
                    return false;
                }
            }
            return true;
        }

        private static bool ValuesEqual(string[][] expectedValue, JToken actualValue)
        {
            var allButHeaderRow = ConvertJaggedArrayType(GetAllButHeaderRow());
            var value = actualValue.ToObject<JArray>().ToObject<string[][]>()[0];
            for (int i = 0; i < expectedValue.Length; i++)
            {
                if (expectedValue[0][i] != value[i])
                {
                    return false;
                }
            }
            return true;
        }

        private static bool JTokenEqualsJaggedArray(JToken token, string[][] rows)
        {
            string[,] stringMultiArray = token.ToObject<JArray>().ToObject<string[,]>();
            int index = 0;
            foreach (var row in rows)
            {
                if (!string.Equals(stringMultiArray[index, 0], row[0]) || !string.Equals(stringMultiArray[index, 1], row[1]))
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

        private static string GetRangeAsNodeNestedArray()
        {
            return "[[\"Name1\", \"Value1\"],[\"Name2\", \"Value2\"]]";
        }

        private static void ResetState()
        {
            finalTable = null;
            finalRange = null;
            finalRangePocoArray = null;
        }

        private class ExcelInputFunctions
        {
            public void GetWorkbookTable([Excel(
                Path = path,
                TableName = tableName)] WorkbookTable table)
            {
                finalTable = table;
            }

            public void GetExcelTableRange(
                [Excel(
                Path = path,
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

            public void AppendRowPocoArray(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] ICollector<SamplePoco> rows)
            {
                foreach (var row in GetRangeAsPocoList())
                {
                    rows.Add(row);
                }
            }

            public void AppendRowNode(
                [Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out string rows)
            {
                rows = GetRangeAsNodeNestedArray();
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

            public void UpdateWorksheetJaggedArray([Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName,
                UpdateType = "Update")] out object[][] rows)
            {
                rows = ConvertJaggedArrayType(GetAllButHeaderRow());
            }

            public void UpdateWorksheetNode([Excel(
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName,
                UpdateType = "Update")] out string rows)
            {
                rows = GetRangeAsNodeNestedArray();
            }
        }

        private class SamplePoco
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }
    }
}
