using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Application.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Data;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CreateSheetWithHeadAndKeyConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(CreateSheetWithHeadAndKeyConsoleApp)}\n");

            try
            {
                // Main class for interacting with Google Sheets API
                var app = new GCPApplication();

                // To use the application you need to authorize the user.
                // For the creating a sheet, you need to authorize a human user account, not service account.
                app.AuthenticateAs(new UserAccount(Properties.Resources.credentials, OAuthSheetsScope.FullAccess));

                Console.WriteLine("Authenticate completed");

                // This uri can be used by everyone because the table is open to everyone.
                const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";
                string[] head = { "title_1", "title_2", "title_3", "title_4" };

                SheetModel sheet = app.CreateSheetWithHeadAndKey(
                    HttpUtils.GetSpreadsheetIdFromUri(uri),
                    "New sheeeeeet123345345",
                    head,
                    head[1]
                );
                PrintSheet(sheet, "Original sheet");

                ChangeSheet(sheet);
                PrintSheet(sheet, "Changed sheet before updating to google");

                // In order for the data in the google spreadsheet sheet to be updated,
                // you need to pass the changed instance of the SheetModel to the UpdateSheet method
                app.UpdateSheet(sheet);
                PrintSheet(sheet, "Sheet after update in google");
            }
            #region User Exceptions
            catch (AuthenticationTimedOutException)
            {
                Console.WriteLine(
                    "Authentication is timed out.\n" +
                    "Run the plugin again and authenticate in the browser."
                );
            }
            catch (UserCanceledAuthenticationException)
            {
                Console.WriteLine(
                    "You have canceled authentication.\n" +
                    $"You need to be authenticated to use library {nameof(SynSys.GSpreadsheetEasyAccess)}."
                );
            }
            catch (UserAccessDeniedException e)
            {
                Console.WriteLine(
                    $"User is denied access to the operation: {e.Operation}\n" +
                    $"Reason: {e.Message}\n" +
                    "Check table access.\n" +
                    "The table must be available to all users who have the link."
                 );
            }
            #endregion
            #region Sheets Exceptions
            catch (SpreadsheetNotFoundException e)
            {
                Console.WriteLine(
                    $"Failed to get spreadsheet with\n" +
                    $"spreadsheet id: {e.SpreadsheetId}\n" +
                    "Check if this spreadsheet exists."
                );
            }
            catch (SheetKeyNotFoundException e)
            {
                Console.WriteLine(
                    $"Key column \"{e.Sheet.KeyName}\" require in the sheet\n" +
                    $"spreadsheet: \"{e.Sheet.SpreadsheetTitle}\"\n" +
                    $"sheet: \"{e.Sheet.Title}\"\n" +
                    "for correct plugin operation\n"
                );
            }
            catch (InvalidSheetHeadException e)
            {
                Console.WriteLine(
                    $"Spreadsheet \"{e.Sheet.SpreadsheetTitle}\"\n" +
                    $"sheet \"{e.Sheet.Title}\"\n" +
                    "lacks required headers:\n" +
                    $"{string.Join(";\n", e.LostHeaders)}."
                );
            }
            catch (SheetExistsException e)
            {
                Console.WriteLine(
                    $"spreadsheet id: \"{e.SpreadsheetId}\"\n" +
                    $"spreadsheet: \"{e.SpreadsheetTitle}\"\n" +
                    $"sheet: \"{e.SheetTitle}\"\n" +
                    "already exists."
                );
            }
            #endregion
            catch (Exception e)
            {
                Console.WriteLine(
                    "An unhandled exception occurred.\n" +
                    $"{e}"
                );
            }

            Console.ReadLine();
        }


        private static void ChangeSheet(SheetModel sheet)
        {
            AddRows(sheet);
            ChangeRows(sheet);
        }

        private static void AddRows(SheetModel sheet)
        {
            // Add empty row
            sheet.AddRow();

            // Add a row part of which will be filled with empty cells
            sheet.AddRow(new string[] { "123", "asd" });

            // Add a row where part of the data will not fall into the line, namely "e" to "l",
            // because test sheet has 4 columns.
            sheet.AddRow(new string[] { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l" });
        }

        private static void ChangeRows(SheetModel sheet)
        {
            // This example doesn't take into account the absence of a key with the desired value
            // and a cell with the selected Title

            // Change added rows with status RowStatus.ToAppend
            sheet.Rows.Last().Cells[0].Value = "change";
            sheet.Rows.Last().Cells[1].Value = "added";
            sheet.Rows.Last().Cells[2].Value = "line";
        }

        private static void PrintSheet(SheetModel sheet, string status)
        {
            PrintDescription(sheet, status);
            PrintHead(sheet.Head);
            PrintBody(sheet.Rows);
        }

        private static void PrintDescription(SheetModel sheet, string status)
        {
            Console.WriteLine(
                "\n" +
               $"Status:            {status}\n" +
               $"Spreadsheet Title: {sheet.SpreadsheetTitle}\n" +
               $"Sheet Title:       {sheet.Title}\n" +
               $"Number of lines:   {sheet.Rows.Count}\n"
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
