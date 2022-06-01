using Newtonsoft.Json;
using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Data;
using System;
using System.Collections.Generic;

namespace SheetMergeConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(SheetMergeConsoleApp)}\n");

            // Main class for interacting with Google Sheets API
            var app = new GCPApplication();

            // To use the application you need to authorize the user.
            app.AuthenticateAs(new UserAccount(Properties.Resources.credentials, OAuthSheetsScope.FullAccess));

            Console.WriteLine("Authorize completed");

            // For tests, you need to change the given uri to your own.
            // Don't forget change keyName on GetSheetWithHeadAndKey and ChangeSheet method
            const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0";

            SheetModel sheetForChange = app.GetSheetWithHeadAndKey(uri, "Head 1");

            // Create a duplicate of the original sheet
            SheetModel originalSheet = JsonSerialization.DeserializeSheet(
                JsonSerialization.SerializeSheet(sheetForChange, Formatting.None)
            );
            PrintSheet(originalSheet, "Original sheet");

            // Add some changes
            ChangeSheet(sheetForChange);

            // Making changes to the original sheet by merging with its modified version
            originalSheet.Merge(sheetForChange);
            PrintSheet(originalSheet, "Merged sheet before updating to google");

            app.UpdateSheet(originalSheet);
            PrintSheet(originalSheet, "Sheet after update in google");

            Console.ReadLine();
        }


        private static void ChangeSheet(SheetModel sheet)
        {
            sheet.AddRow(new List<string>() { "a", "b", "c", "d", "e" });
            sheet.DeleteRow(sheet.Rows[5]);
            sheet.Rows[1].Cells[1].Value = "22222";
        }

        private static void PrintSheet(SheetModel sheet, string status)
        {
            PrintDesctiption(sheet, status);
            PrintHead(sheet.Head);
            PrintBody(sheet.Rows);
        }

        private static void PrintDesctiption(SheetModel sheet, string status)
        {
            Console.WriteLine(
                "\n" +
               $"Status:           {status}\n" +
               $"Spreadsheet Name: {sheet.SpreadsheetTitle}\n" +
               $"Sheet Name:       {sheet.Title}\n" +
               $"Number of lines:  {sheet.Rows.Count}\n"
            );
        }

        private static void PrintHead(List<string> head)
        {
            Console.Write($"{"", 3}");

            foreach (string title in head)
            {
                string value = title;

                if (string.IsNullOrWhiteSpace(title))
                {
                    value = "-";
                }

                Console.Write($"|{value,7}");
            }

            Console.Write($"| ");
            Console.WriteLine();

            Console.Write($"{"", 3}");
            var delimiter = new String('-', 7);

            foreach (string title in head)
            {
                Console.Write($"|{delimiter}");
            }

            Console.Write($"| ");
            Console.WriteLine();
        }

        private static void PrintBody(List<Row> rows)
        {
            foreach (var row in rows)
            {
                Console.Write($"{row.Number,3}");

                foreach (var cell in row.Cells)
                {
                    Console.Write($"|{cell.Value,7}");
                }

                Console.Write($"| {row.Status}");
                Console.WriteLine();
            }
        }
    }
}
