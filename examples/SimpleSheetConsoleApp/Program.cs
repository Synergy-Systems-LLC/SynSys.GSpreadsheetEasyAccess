using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Data;
using System;
using System.Collections.Generic;

namespace SimpleSheetConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(SimpleSheetConsoleApp)}\n");

            // Main class for interacting with Google Sheets API
            var app = new GCPApplication();

            // To use the application you need to authorize the user.
            app.AuthenticateAs(new ServiceAccount(Properties.Resources.apikey));

            Console.WriteLine("Authorize completed");

            // This uri can be used by everyone because the table is open to everyone.
            const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";

            SheetModel sheet = app.GetSheet(uri);
            PrintSheet(sheet, "Original sheet");

            Console.ReadLine();
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
