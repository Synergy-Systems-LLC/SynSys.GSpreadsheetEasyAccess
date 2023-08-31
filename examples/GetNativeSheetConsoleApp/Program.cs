using Printing;
using SynSys.GSpreadsheetEasyAccess.Application;
using SynSys.GSpreadsheetEasyAccess.Application.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Data;
using System;

namespace GetNativeSheetConsoleApp
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine($"Start {nameof(GetNativeSheetConsoleApp)}\n");

            try
            {
                // Main class for interacting with Google Sheets API
                var app = new GCPApplication();

                //To use the application you need to authorize the user.
                //app.AuthenticateAs(new ServiceAccount(Properties.Resources.apikey));
                app.AuthenticateAs(new UserAccount(Properties.Resources.credentials, OAuthSheetsScope.FullAccess));

                Console.WriteLine("Authenticate completed");

                // This uri can be used by everyone because the table is open to everyone.
                const string uri = "https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0&fvid=545275384";

                NativeSheet sheet = app.GetNativeSheet(new GoogleSheetUri(uri));
                Printer.PrintSheet(sheet, "Original sheet");

                // After adding a column, the statuses of the rows do not change
                sheet.AddColumn();
                Printer.PrintSheet(sheet, "Sheet with new added column");

                app.UpdateSheet(sheet);
                Printer.PrintSheet(sheet, "Sheet after update in google");

                // After inserting a column, the statuses of the rows do not change
                sheet.InsertColumn(4);
                sheet.InsertColumn(5);
                sheet.InsertColumn(7);
                sheet.AddColumn();
                Printer.PrintSheet(sheet, "Sheet with new inserted column");

                app.UpdateSheet(sheet);
                Printer.PrintSheet(sheet, "Sheet after update in google");
            }
            #region User Exceptions
            catch (InvalidApiKeyException)
            {
                Console.WriteLine(
                    "There was a problem with the google service api key. Check its validity"
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
    }
}
