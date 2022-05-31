using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SheetWithHeadAndKeyConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(SheetWithHeadAndKeyConsoleApp)}\n");

            // Main class for interacting with Google Sheets API
            var app = new GCPApplication();

            // To use the application you need to authorize the user
            app.AuthenticateAs(new UserAccount(Properties.Resources.credentials, OAuthSheetsScope.FullAccess));

            Console.WriteLine("Authorize completed");

            // For tests, you need to change the given uri to your own.
            // Don't forget change keyName on GetSheetWithHeadAndKey and ChangeSheet method
            const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";

            SheetModel sheet = app.GetSheetWithHeadAndKey(uri, "Head 1");
            PrintSheet(sheet, "Original sheet");

            ChangeSheet(sheet);
            PrintSheet(sheet, "Changed sheet before updating to google");

            // In order for the data in the google spreadsheet sheet to be updated,
            // you need to pass the changed instance of the SheetModel to the UpdateSheet method
            app.UpdateSheet(sheet);
            PrintSheet(sheet, "Sheet after update in google");

            // The function is called a second time to show
            // that the current instance of the sheet can still be used after the update.
            // Its structure will match the google spreadsheet.
            ChangeSheet(sheet);
            PrintSheet(sheet, "Changed sheet before updating to google");

            app.UpdateSheet(sheet);
            PrintSheet(sheet, "Sheet after update in google");

            Console.ReadLine();
        }


        private static void ChangeSheet(SheetModel sheet)
        {
            AddRows(sheet);
            ChangeRows(sheet);
            DeleteRows(sheet);
        }

        private static void AddRows(SheetModel sheet)
        {
            // Add empty row
            sheet.AddRow();

            // Add a row part of which will be filled with empty cells
            sheet.AddRow(new List<string>() { "123", "asd" });

            // Add a row where part of the data will not fall into the line, namely "k" and "l",
            // because test sheet has 10 columns.
            sheet.AddRow(
                new List<string>()
                {
                    "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l"
                }
            );
        }

        private static void ChangeRows(SheetModel sheet)
        {
            // This example doesn't take into account the absence of a key with the desired value
            // and a cell with the selected Title
            sheet.Rows.Find(row => row.Key.Value == "31")
                 .Cells.Find(cell => cell.Title == "Head 6")
                 .Value = "360";
            sheet.Rows.Find(row => row.Key.Value == "61")
                 .Cells.Find(cell => cell.Title == "Head 3")
                 .Value = "630";
            sheet.Rows.Find(row => row.Key.Value == "51")
                 .Cells.Find(cell => cell.Title == "Head 2")
                 .Value = "520";

            // Change added rows with status RowStatus.ToAppend
            sheet.Rows.Last().Cells[0].Value = "change";
            sheet.Rows.Last().Cells[1].Value = "added";
            sheet.Rows.Last().Cells[2].Value = "line";
        }

        private static void DeleteRows(SheetModel sheet)
        {
            // This example doesn't take into account the lack of necessary indexes
            sheet.DeleteRow(sheet.Rows[3]);

            // Delete the added line with the status RowStatus.ToAppend.
            // In this case, the line is immediately deleted without waiting for the sheet to be updated.
            sheet.DeleteRow(sheet.Rows[10]);
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
