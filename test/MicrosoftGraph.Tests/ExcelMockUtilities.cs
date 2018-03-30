// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

namespace Microsoft.Azure.WebJobs.Extensions.MicrosoftGraph.Tests
{
    using System;
    using System.Linq.Expressions;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Moq;
    using Newtonsoft.Json.Linq;

    public static class ExcelMockUtilities
    {
        public static void MockGetTableWorkbookAsync(this Mock<IGraphServiceClient> mock, WorkbookTable returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Tables[It.IsAny<string>()]
                .Request()
                .GetAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void MockGetTableWorkbookRangeAsync(this Mock<IGraphServiceClient> mock, WorkbookRange returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Tables[It.IsAny<string>()]
                .Range()
                .Request(null)
                .GetAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void MockGetWorksheetWorkbookAsync(this Mock<IGraphServiceClient> mock, WorkbookRange returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Worksheets[It.IsAny<string>()]
                .UsedRange()
                .Request(null)
                .GetAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void MockGetWorkSheetWorkbookInRangeAsync(this Mock<IGraphServiceClient> mock, WorkbookRange returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Worksheets[It.IsAny<string>()]
                .Range(It.IsAny<string>())
                .Request(null)
                .GetAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void MockGetTableHeaderRowAsync(this Mock<IGraphServiceClient> mock, WorkbookRange returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Tables[It.IsAny<string>()]
                .HeaderRowRange()
                .Request(null)
                .GetAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void MockPostTableRowAsyc(this Mock<IGraphServiceClient> mock, WorkbookTableRow returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Tables[It.IsAny<string>()]
                .Rows
                .Add(null, It.IsAny<JToken>())
                .Request(null)
                .PostAsync()).Returns(Task.FromResult(returnValue));
        }

        public static void VerifyPostTableRowAsync(this Mock<IGraphServiceClient> mock, string path, string tableName, Expression<Func<JToken,bool>> rowCondition)
        {
            //first verify PostAsync() called
            mock.Verify(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Tables[tableName]
                .Rows
                .Add(null, It.IsAny<JToken>())
                .Request(null)
                .PostAsync());

            //Now verify row condition is true
            mock.Verify(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Tables[tableName]
                .Rows
                .Add(null, It.Is<JToken>(rowCondition)));
        }

        public static void MockPatchWorksheetAsync(this Mock<IGraphServiceClient> mock, WorkbookRange returnValue)
        {
            mock.Setup(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(It.IsAny<string>())
                .Workbook
                .Worksheets[It.IsAny<string>()]
                .Range(It.IsAny<string>())
                .Request(null)
                .PatchAsync(It.IsAny<WorkbookRange>())).Returns(Task.FromResult(returnValue));
        }

        public static void VerifyPatchWorksheetAsync(this Mock<IGraphServiceClient> mock, string path, string worksheetName, string range, Expression<Func<WorkbookRange,bool>> newWorkbookCondition)
        {
            mock.Verify(client => client
                .Me
                .Drive
                .Root
                .ItemWithPath(path)
                .Workbook
                .Worksheets[worksheetName]
                .Range(range)
                .Request(null)
                .PatchAsync(It.Is<WorkbookRange>(newWorkbookCondition)));
        }
    }
}
