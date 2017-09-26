// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System.Threading.Tasks;
    using Xunit;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.Token.Tests;
    using Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Services;
    using Moq;
    using Microsoft.Graph;
    using System.Collections.Generic;
    using Newtonsoft.Json.Linq;
    using System.Linq;
    using Newtonsoft.Json;

    public class ExcelTestsEndToEnd
    {
        private static WorkbookTable finalTable;
        private static string[][] finalRange;
        private static SamplePoco[] finalRangePocoArray;
        private static List<SamplePoco> finalRangePocoList;
        private static object[][] finalRangeOut;

        private const string path = "sample/path";
        private const string tableName = "tableName";
        private const string worksheetName = "worksheetName";
        private const string fullOldTableAddress = "worksheetName!A1:B5";
        private const string newTableAddressMinusHeader = "worksheetName!A2:B3";

        [Fact]
        public static async Task GetWorkbookTable()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            WorkbookTable table = new WorkbookTable();
            excelMock.Setup(client => client.GetTableWorkbookAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(table));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelInputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelInputFunctions.GetWorkbookTable", args);
            var expectedResult = table;
            Assert.Equal(expectedResult, finalTable);
            ResetState();
        }

        [Fact]
        public static async Task GetWorkbookTableAsJaggedStringArray()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[][] range = GetRangeAsJaggedStringArray();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(range)
            };
            excelMock.Setup(client => client.GetTableWorkbookRangeAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(workbookRange));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelInputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelInputFunctions.GetExcelTableRange", args);
            var expectedResult = range;
            Assert.Equal(expectedResult, finalRange);
            ResetState();
        }

        [Fact]
        public static async Task GetWorksheetRangeAsJaggedStringArray()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[][] range = GetRangeAsJaggedStringArray();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(range)
            };
            excelMock.Setup(client => client.GetWorksheetWorkbookAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(workbookRange));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelInputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelInputFunctions.GetExcelWorksheetRange", args);
            var expectedResult = range;
            Assert.Equal(expectedResult, finalRange);
            ResetState();
        }

        [Fact]
        public static async Task GetWorksheetRangeAsPocoArray()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[][] range = GetRangeAsJaggedStringArray();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(range)
            };
            excelMock.Setup(client => client.GetWorksheetWorkbookAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(workbookRange));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelInputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelInputFunctions.GetExcelWorksheetRangePocoArray", args);
            var expectedResult = JsonConvert.SerializeObject(GetRangeAsPocoArray());
            Assert.Equal(expectedResult, JsonConvert.SerializeObject(finalRangePocoArray));
            ResetState();
        }

        [Fact]
        public static async Task GetWorksheetRangeAsPocoList()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[][] range = GetRangeAsJaggedStringArray();
            var workbookRange = new WorkbookRange()
            {
                Values = JToken.FromObject(range)
            };
            excelMock.Setup(client => client.GetWorksheetWorkbookAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(workbookRange));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelInputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelInputFunctions.GetExcelWorksheetRangePocoList", args);
            var expectedResult = JsonConvert.SerializeObject(GetRangeAsPocoList());
            Assert.Equal(expectedResult, JsonConvert.SerializeObject(finalRangePocoList));
            ResetState();
        }

        [Fact]
        public static async Task AppendRowsWithStringJaggedArray()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[] headerRow = GetHeaderRow();
            excelMock.Setup(client => client.GetTableHeaderRowAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(headerRow));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelOutputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelOutputFunctions.AppendRowJaggedArray", args);
            excelMock.Verify(client => client.PostTableRowAsync(path, tableName, It.Is<JToken>(token => JTokenEqualsAllButHeaderRow(token))));
            ResetState();
        }

        [Fact]
        public static async Task AppendRowWithPoco()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[] headerRow = GetHeaderRow();
            excelMock.Setup(client => client.GetTableHeaderRowAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(headerRow));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelOutputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelOutputFunctions.AppendRowPoco", args);
            var samplePoco = GetRangeAsPocoArray()[0];
            excelMock.Verify(client => client.PostTableRowAsync(path, tableName, It.Is<JToken>(token => JTokenEqualsPoco(token, samplePoco))));
            ResetState();
        }

        [Fact]
        public static async Task AppendRowsWithPocoList()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[] headerRow = GetHeaderRow();
            excelMock.Setup(client => client.GetTableHeaderRowAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(headerRow));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelOutputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelOutputFunctions.AppendRowPocoList", args);
            var samplePocos = GetRangeAsPocoList();
            excelMock.Verify(client => client.PostTableRowAsync(path, tableName, It.Is<JToken>(token => JTokenEqualsPocos(token, samplePocos))));
            ResetState();
        }

        [Fact]
        public static async Task AppendRowsWithPocoArray()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[] headerRow = GetHeaderRow();
            excelMock.Setup(client => client.GetTableHeaderRowAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(headerRow));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelOutputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelOutputFunctions.AppendRowPocoArray", args);
            var samplePocos = GetRangeAsPocoArray();
            excelMock.Verify(client => client.PostTableRowAsync(path, tableName, It.Is<JToken>(token => JTokenEqualsPocos(token, samplePocos))));
            ResetState();
        }

        [Fact]
        public static async Task UpdateWorksheetWithTable()
        {
            var graphConfig = new MicrosoftGraphExtensionConfig();
            var excelMock = new Mock<IExcelClient>();
            string[] headerRow = GetHeaderRow();
            var address = new WorkbookRange()
            {
                Address = fullOldTableAddress,
            };
            excelMock.Setup(client => client.GetWorksheetWorkbookAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(address));
            excelMock.Setup(client => client.GetTableHeaderRowAsync(It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(headerRow));
            var workbookInRange = new WorkbookRange();
            excelMock.Setup(client => client.GetWorkSheetWorkbookInRangeAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>())).Returns(Task.FromResult(workbookInRange));
            graphConfig.ExcelClient = excelMock.Object;

            var jobHost = TestHelpers.NewHost<ExcelOutputFunctions>(graphConfig);
            var args = new Dictionary<string, object>();
            await jobHost.CallAsync("ExcelOutputFunctions.UpdateWorksheet", args);
            var samplePocos = GetRangeAsPocoArray();
            excelMock.Verify(client => client.GetWorkSheetWorkbookInRangeAsync(path, worksheetName, newTableAddressMinusHeader));
            excelMock.Verify(client => client.PatchWorksheetAsync(path, worksheetName, newTableAddressMinusHeader, It.Is<WorkbookRange>(range =>JTokenEqualsPocos(range.Values, samplePocos))));
            ResetState();
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

        private static bool JTokenEqualsPoco(JToken token, SamplePoco poco)
        {
            string[,] stringMultiArray = token.ToObject<JArray>().ToObject<string[,]>();
            return string.Equals(stringMultiArray[0, 0], poco.Name) && string.Equals(stringMultiArray[0, 1], poco.Value);
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
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] WorkbookTable table)
            {
                finalTable = table;
            }

            public void GetExcelTableRange(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] string[][] range)
            {
                finalRange = range;
            }

            public void GetExcelWorksheetRange(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName)] string[][] range)
            {
                finalRange = range;
            }

            public void GetExcelWorksheetRangePocoArray(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName)] SamplePoco[] range)
            {
                finalRangePocoArray = range;
            }

            public void GetExcelWorksheetRangePocoList(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
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
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out object[][] range)
            {
                range = ConvertJaggedArrayType(GetAllButHeaderRow());
            }

            public void AppendRowPoco(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out SamplePoco row)
            {
                row = GetRangeAsPocoArray()[0];
            }

            public void AppendRowPocoList(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out List<SamplePoco> rows)
            {
                rows = GetRangeAsPocoList();
            }

            public void AppendRowPocoArray(
                [Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
                Path = path,
                WorksheetName = worksheetName,
                TableName = tableName)] out SamplePoco[] rows)
            {
                rows = GetRangeAsPocoArray();
            }

            public void UpdateWorksheet([Excel(
                UserId = "UserId",
                IdentityProvider = "AAD",
                Identity = TokenIdentityMode.UserFromId,
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
