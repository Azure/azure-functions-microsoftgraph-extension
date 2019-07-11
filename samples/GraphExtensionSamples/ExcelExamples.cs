// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
namespace GraphExtensionSamples
{
    using Microsoft.Azure.WebJobs;
    using Microsoft.AspNetCore.Http;
    using System.Collections.Generic;
    using Microsoft.Graph;

    public static class ExcelExamples
    {
        //Appending row(s) to an excel worksheet or table

        public static void AppendRowsToExcelSpreadsheetWithJaggedArray([Excel(
            Path = "TestSheet.xlsx",
            WorksheetName = "Sheet1",
            Identity = TokenIdentityMode.UserFromRequest)] out object[][] output)
        {
            output = new object[2][];
            output[0] = new object[]
            {
                "samplepartname", 42, 3.75,
            };
            output[1] = new object[]
            {
                "part2", 73, 43.20,
            };
        }

        [NoAutomaticTrigger]
        public static void AppendRowToExcelTableWithSinglePoco([Excel(
            Path = "TestSheet.xlsx",
            WorksheetName = "Sheet1",
            TableName = "Parts",
            Identity = TokenIdentityMode.UserFromRequest)] out PartsTableRow output)
        {
            output = new PartsTableRow
            {
                Part = "samplepartname",
                Id = 42,
                Price = 3.75,
            };
        }

        //Updating an excel table

        [NoAutomaticTrigger]
        public static void UpdateEntireExcelTableWithPocos(
            [Excel(
            Path = "TestSheet.xlsx",
            WorksheetName = "Sheet1",
            TableName = "Parts",
            UpdateType = "Update",
            Identity = TokenIdentityMode.UserFromRequest)] out PartsTableRow[] output)
        {
            output = new PartsTableRow[2];
            output[0] = (new PartsTableRow
            {
                Part = "part1",
                Id = 35,
                Price = 0.75
            });
            output[1] = (new PartsTableRow
            {
                Part = "part2",
                Id = 73,
                Price = 42.37,
            });
        }

        //Retrieving values from an excel table or worksheet
        public static void GetEntireExcelWorksheetAsJaggedStringArray([Excel(
            Path = "TestSheet.xlsx",
            WorksheetName = "Sheet1",
            Identity = TokenIdentityMode.UserFromRequest)] string[][] rows)
        {
            //Perform any operations on the string[][], where each string[] is 
            //a row in the worksheet.
        }

        public static void GetExcelTableAsWorkbookTable([Excel(
            Path = "TestSheet.xlsx",
            WorksheetName = "Sheet1",
            TableName = "sampletable",
            Identity = TokenIdentityMode.UserFromRequest)] WorkbookTable table)
        {
            //See https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/81c50e72166152f9f84dc38b2516379b7a536300/src/Microsoft.Graph/Models/Generated/WorkbookTable.cs
            //for usage
        }

        public static void GetExcelTableAsPocoList([Excel(
            Path = "TestSheet.xlsx",
            WorksheetName = "Sheet1",
            TableName = "sampletable",
            Identity = TokenIdentityMode.UserFromRequest)] PartsTableRow table)
        {
            //Note that each POCO object represents one row, and the values correspond to
            //the column titles that match the POCO's property names.
        }

        public class PartsTableRow
        {
            public string Part { get; set; }

            public int Id { get; set; }

            public double Price { get; set; }
        }
    }
}
