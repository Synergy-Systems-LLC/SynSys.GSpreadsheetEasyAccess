using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Учетные записи служб управляются <a href="https://cloud.google.com/iam/docs/understanding-service-accounts">IAM</a>, 
    /// и они представляют пользователей, не являющихся людьми.<br/>
    /// Они предназначены для сценариев, когда вашему приложению требуется доступ к ресурсам или
    /// выполнение действий самостоятельно, <br/>
    /// например запуск приложений App Engine или взаимодействие с экземплярами Compute Engine.<br/>
    /// Дополнительные сведения см. в разделе 
    /// <a href="https://cloud.google.com/docs/authentication/api-keys">Аутентификация в качестве учетной записи службы</a>.
    /// </summary>
    public class ServiceAccount : Principal
    {
        private string _apiKey;

        /// <summary>
        /// Аутентифицировать приложение как учетную запись службы.
        /// </summary>
        /// <returns></returns>
        public ServiceAccount(string apikey)
        {
            _apiKey = apikey;
        }

        /// <summary>
        /// Вернёт объект представляющий сервис таблиц.<br/>
        /// Это 
        /// <a href="https://developers.google.com/resources/api-libraries/documentation/sheets/v4/csharp/latest/classGoogle_1_1Apis_1_1Sheets_1_1v4_1_1SheetsService.html">
        /// объект библиотеки Google.Apis.Sheets.v4
        /// </a>,
        /// через который происходит работа с таблицами Google.
        /// </summary>
        /// <returns></returns>
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
