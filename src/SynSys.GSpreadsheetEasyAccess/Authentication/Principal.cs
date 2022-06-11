using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Principal is an object that can be granted access to a resource.<br/>
    /// See <a href="https://cloud.google.com/docs/authentication">Authentication Overview</a> for details.
    /// </summary>
    public abstract class Principal
    {
        /// <summary>
        /// Return an object representing Google Sheets service.<br/>
        /// </summary>
        /// <remarks>
        /// It
        /// <a href="https://developers.google.com/resources/api-libraries/documentation/sheets/v4/csharp/latest/classGoogle_1_1Apis_1_1Sheets_1_1v4_1_1SheetsService.html">
        /// Google.Apis.Sheets.v4 library object
        /// </a>,
        /// through which work with Google Sheets API takes place.
        /// </remarks>
        public abstract SheetsService GetSheetsService();
    }
}
