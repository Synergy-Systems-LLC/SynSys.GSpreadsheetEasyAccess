using Newtonsoft.Json;
using Printing;
using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Application.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Data;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;
using System;
using System.Collections.Generic;

namespace SheetMergeConsoleApp
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(SheetMergeConsoleApp)}\n");

            try
            {
                // Main class for interacting with Google Sheets API
                var app = new GCPApplication();

                // To use the application you need to authorize the user.
                app.AuthenticateAs(new UserAccount(Properties.Resources.credentials, OAuthSheetsScope.FullAccess));

                Console.WriteLine("Authorize completed");

                // For tests, you need to change the given uri to your own.
                // Don't forget change keyName on GetSheetWithHeadAndKey and ChangeSheet method
                const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0";

                UserSheet originalSheet = app.GetUserSheet(new GoogleSheetUri(uri), "Title 1");
                Printer.PrintSheet(originalSheet, "Original sheet");

                // Create a duplicate of the original sheet
                UserSheet sheetForChange = JsonSerialization.DeserializeSheet<UserSheet>(
                    JsonSerialization.SerializeSheet(originalSheet, Formatting.None)
                );

                // Add some changes
                ChangeSheet(sheetForChange);
                Printer.PrintSheet(sheetForChange, "Changed sheet");

                // Making changes to the original sheet by merging with its modified version
                originalSheet.Merge(sheetForChange);
                Printer.PrintSheet(originalSheet, "Merged sheet before updating on google");

                app.UpdateSheet(originalSheet);
                Printer.PrintSheet(originalSheet, "Sheet after update in google");
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
            sheet.AddRow(new List<string>() { "a", "b", "c", "d", "e" });
            sheet.DeleteRow(sheet.Rows[5]);
            sheet.Rows[1].Cells[1].Value = "22222";
        }
    }
}
