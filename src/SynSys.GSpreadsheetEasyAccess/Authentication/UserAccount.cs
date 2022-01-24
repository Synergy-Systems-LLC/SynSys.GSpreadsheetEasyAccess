using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Reflection;
using System.Threading;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Учетные записи пользователей управляются как учетные записи Google, 
    /// они представляют разработчика, администратора или любого другого человека, который взаимодействует с Google Cloud.<br/>
    /// Они предназначены для сценариев, в которых вашему приложению требуется доступ к ресурсам от имени пользователя-человека.<br/>
    /// Дополнительную информацию см. в разделе
    /// <a href="https://cloud.google.com/docs/authentication/end-user">Аутентификация в качестве конечного пользователя</a>.
    /// </summary>
    public class UserAccount : Principal
    {
        private byte[] _credentials;
        private OAuthSheetsScope _scope;

        /// <summary>
        /// Аутентификации в качестве пользователя-человека.
        /// </summary>
        /// <returns></returns>
        public UserAccount(byte[] credentials, OAuthSheetsScope scope)
        {
            _credentials = credentials;
            _scope = scope;
        }

        /// <summary>
        /// Время на попытку подключения к приложению на Google Cloud Platform в секундах.<br/>
        /// Значение по умолчанию <c>15</c>. 
        /// </summary>
        public byte CancellationSeconds { get; set; } = 15;

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
                    HttpClientInitializer = GetUserCredential(new MemoryStream(_credentials))
                }
            );
        }


        private UserCredential GetUserCredential(Stream credentials)
        {
            return GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(credentials).Secrets,
                _scope.Value,
                "user",
                GenerateCancellationToken(),
                new FileDataStore(GenerateTokenPath(), true)
            ).Result;
        }

        private CancellationToken GenerateCancellationToken()
        {
            var tokenSource = new CancellationTokenSource();
            tokenSource.CancelAfter(TimeSpan.FromSeconds(CancellationSeconds));
            return tokenSource.Token;
        }

        private string GenerateTokenPath()
        {
            var uriPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().EscapedCodeBase);
            var outputDirectory = new Uri(uriPath).LocalPath;
            return Path.Combine(outputDirectory, "token.json");
        }
    }
}
