using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Service accounts are managed by <a href="https://cloud.google.com/iam/docs/understanding-service-accounts">IAM</a>,
    /// and represent non-human users.
    /// </summary>
    /// <remarks>
    /// These are for scenarios where your application needs to access resources or
    /// perform actions on their own, <br/>
    /// such as running App Engine applications or interacting with Compute Engine instances.<br/>
    /// For more information, see
    /// <a href="https://cloud.google.com/docs/authentication/api-keys">Authenticate as a service account</a>.
    /// </remarks>
    public class ServiceAccount : Principal
    {
        private string _apiKey;

        /// <summary>
        /// Authenticate the application as a service account.
        /// </summary>
        /// <returns></returns>
        public ServiceAccount(string apikey)
        {
            _apiKey = apikey;
        }

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
        public override SheetsService GetSheetsService()
        {
            return new SheetsService(
                new BaseClientService.Initializer
                {
                    ApiKey = _apiKey
                }
            );
        }
    }
}
