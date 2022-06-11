using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// OAuth scopes restrict the actions your application can perform on behalf of the end user.<br/>
    /// Set the minimum size required for your use case.<br/>
    /// The class implements only some of the areas from
    /// <a href="https://developers.google.com/identity/protocols/oauth2/scopes#sheets">Google Sheets API v4</a>
    /// </summary>
    public class OAuthSheetsScope
    {
        internal OAuthSheetsScope(string[] scopes)
        {
            Value = scopes;
        }

        internal string[] Value { get; set; }

        /// <summary>
        /// Access to view Google spreadsheet.
        /// </summary>
        public static OAuthSheetsScope ViewAccess => new OAuthSheetsScope(
            new string[] { SheetsService.Scope.SpreadsheetsReadonly }
        );

        /// <summary>
        /// Access to view, edit, create and delete Google spreadsheet.
        /// </summary>
        public static OAuthSheetsScope FullAccess => new OAuthSheetsScope(
            new string[] { SheetsService.Scope.Spreadsheets }
        );
    }
}
