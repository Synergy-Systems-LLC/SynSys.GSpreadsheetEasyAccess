﻿using Newtonsoft.Json;
using Printing;
using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Application.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Data;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SerializationSheetModelConsoleApp
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(SerializationSheetModelConsoleApp)}\n");

            try
            {
                // Main class for interacting with Google Sheets API
                var app = new GCPApplication();

                // To use the application you need to authorize the user.
                app.AuthenticateAs(new UserAccount(Properties.Resources.credentials, OAuthSheetsScope.FullAccess));

                Console.WriteLine("Authorize completed");

                // For tests, you need to change the given uri to your own.
                // Don't forget change keyName on GetSheetWithHeadAndKey and ChangeSheet method
                const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";

                UserSheet sheet = app.GetUserSheet(new GoogleSheetUri(uri), "Title 1");
                Printer.PrintSheet(sheet, "Original sheet");

                // Serialize UserSheet to string json structure
                string jsonSheet = JsonSerialization.SerializeSheet(sheet, Formatting.Indented);
                Console.WriteLine($"\nSheet after serialization\n{jsonSheet}");

                // Deserialize string json structure to SheetModel instance
                var deserializedSheet = JsonSerialization.DeserializeSheet<UserSheet>(jsonSheet);
                Printer.PrintSheet(deserializedSheet, "Sheet after deserialization");

                // Sheet instance can be used after deserialization
                ChangeSheet(deserializedSheet);
                Printer.PrintSheet(deserializedSheet, "Changed sheet before updating to google");
                app.UpdateSheet(deserializedSheet);
                Printer.PrintSheet(deserializedSheet, "Sheet after update in google");
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
                    "Check if this table exists."
                );
            }
            catch (SheetNotFoundException e)
            {
                Console.WriteLine(
                    "Failed to get sheet with\n" +
                    $"spreadsheet id: {e.SpreadsheetId}\n" +
                    $"spreadsheet title: {e.SpreadsheetTitle}\n" +
                    $"sheet id: {e.SheetGid}\n" +
                    "Check if this sheet exists."
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
            catch (EmptySheetException e)
            {
                Console.WriteLine(
                    "For the plugin to work correctly sheet" +
                    $"spreadsheet: \"{e.Sheet.SpreadsheetTitle}\"\n" +
                    $"sheet: \"{e.Sheet.Title}\"\n" +
                    "cannot be empty."
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

        private static void ChangeSheet(AbstractSheet sheet)
        {
            AddRows(sheet);
            ChangeRows(sheet);
            DeleteRows(sheet);
        }

        private static void AddRows(AbstractSheet sheet)
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

        private static void ChangeRows(AbstractSheet sheet)
        {
            // This example doesn't take into account the absence of a key with the desired value
            // and a cell with the selected Title
            sheet.Rows.Find(row => row.Key.Value == "31")
                 .Cells.Find(cell => cell.Column.Title == "Title 6")
                 .Value = "360";
            sheet.Rows.Find(row => row.Key.Value == "61")
                 .Cells.Find(cell => cell.Column.Title == "Title 3")
                 .Value = "630";
            sheet.Rows.Find(row => row.Key.Value == "51")
                 .Cells.Find(cell => cell.Column.Title == "Title 2")
                 .Value = "520";

            // Change added rows with status RowStatus.ToAppend
            sheet.Rows.Last().Cells[0].Value = "change";
            sheet.Rows.Last().Cells[1].Value = "added";
            sheet.Rows.Last().Cells[2].Value = "line";
        }

        private static void DeleteRows(AbstractSheet sheet)
        {
            // This example doesn't take into account the lack of necessary indexes
            sheet.DeleteRow(sheet.Rows[3]);

            // Delete the added line with the status RowStatus.ToAppend.
            // In this case, the line is immediately deleted without waiting for the sheet to be updated.
            sheet.DeleteRow(sheet.Rows[10]);
        }
    }
}
