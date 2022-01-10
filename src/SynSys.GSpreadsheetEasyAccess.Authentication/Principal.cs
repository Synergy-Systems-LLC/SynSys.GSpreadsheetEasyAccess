using Google.Apis.Sheets.v4;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// Principal - это объект которому может быть предоставлен доступ к ресурсу.<br/>
    /// Подробную информацию см. в разделе <a href="https://cloud.google.com/docs/authentication">Обзор аутентификации</a>.
    /// </summary>
    public abstract class Principal
    {
        public abstract SheetsService GetSheetsService();
    }
}
