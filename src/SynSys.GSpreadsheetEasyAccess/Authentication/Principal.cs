using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Principal - это объект которому может быть предоставлен доступ к ресурсу.<br/>
    /// Подробную информацию см. в разделе <a href="https://cloud.google.com/docs/authentication">Обзор аутентификации</a>.
    /// </summary>
    public abstract class Principal
    {
        /// <summary>
        /// Вернёт объект представляющий сервис таблиц.<br/>
        /// Это 
        /// <a href="https://developers.google.com/resources/api-libraries/documentation/sheets/v4/csharp/latest/classGoogle_1_1Apis_1_1Sheets_1_1v4_1_1SheetsService.html">
        /// объект библиотеки Google.Apis.Sheets.v4
        /// </a>,
        /// через который происходит работа с таблицами Google.
        /// </summary>
        /// <returns></returns>
        public abstract SheetsService GetSheetsService();
    }
}
