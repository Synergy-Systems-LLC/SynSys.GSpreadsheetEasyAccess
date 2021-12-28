using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Области действия OAuth ограничивают действия, которые ваше приложение может выполнять от имени конечного пользователя.<br/>
    /// Установите минимальный объем, необходимый для вашего варианта использования.<br/>
    /// В классе реализована только чать областей из
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
        /// Доступ на просмотр таблиц Google.
        /// </summary>
        public static OAuthSheetsScope ViewAccess => new OAuthSheetsScope(
            new string[] { SheetsService.Scope.SpreadsheetsReadonly }
        );

        /// <summary>
        /// Доступ на просмотр, редактирование, создание и удаление таблиц Google.
        /// </summary>
        public static OAuthSheetsScope FullAccess => new OAuthSheetsScope(
            new string[] { SheetsService.Scope.Spreadsheets }
        );
    }
}
