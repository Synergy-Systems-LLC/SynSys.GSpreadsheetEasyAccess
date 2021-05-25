using System;
using System.Text.RegularExpressions;

namespace SynSys.GSpreadsheetEasyAccess
{
    // Сделать данный класс не статичным!
    public static class HttpManager
    {
        public static string Status { get; set; }

        public static string GetSpreadsheetIdFromUri(string uri)
        {
            var match = Regex.Match(uri, @"(?s)(?<=d/).+?(?=/edit)");
            return match.ToString();
        }

        public static int GetGidFromUri(string uri)
        {
            var match = Regex.Match(uri, @"(?s)(?<=gid=)\d+");

            if (string.IsNullOrWhiteSpace(match.Value))
            {
                return -1;
            }

            return int.Parse(match.Value);
        }


        internal static string GetHttpStatusCode(string errorMessage)
        {
            string pattern = @"\[(\d+)\]";
            var rx = new Regex(pattern);

            if (rx.IsMatch(errorMessage))
            {
                return rx.Match(errorMessage).Groups[1].Value;
            }
            else
            {
                return string.Empty;
            }
        }

        internal static string GetMessageByCode(string httpCode)
        {
            switch (httpCode)
            {
                case "400":
                    return "Неверный запрос. Url адрес не существует.";
                case "401":
                    return "Проблемы с авторизацией в google.";
                case "403":
                    return "Отказ в доступе. Обратитесь к владельцу таблицы google.";
                case "404":
                    return "Не найден запрашиваемый адреc. Возможно вы ошиблись в Id таблицы или Id её листа.";
                default:
                    return $"http status code - {httpCode}.";
            }
        }

        internal static bool IsCorrectUrl(string uri)
        {
            if (string.IsNullOrEmpty(uri))
            {
                Status = "Url пустой";
                return false;
            }

            if (NotExistsSpreadsheet(uri))
            {
                Status = "Url адрес не принадлежит google spreadsheet";
                return false;
            }

            if (NotExistsId(uri))
            {
                Status = "В Url адресе некорректный spreadsheetId";
                return false;
            }

            if (NotExistsGid(uri))
            {
                Status = "В Url адресе некорректный gid листа или он отсутствует";
                return false;
            }

            return true;
        }


        private static bool NotExistsSpreadsheet(string url)
        {
            if (url.Contains("spreadsheets"))
            {
                return false;
            }

            return true;
        }

        private static bool NotExistsId(string url)
        {
            var match = Regex.Match(url, @"(?s)(?<=d/).+?(?=/edit)");
            return !match.Success;
        }

        private static bool NotExistsGid(string url)
        {
            var match = Regex.Match(url, @"(?s)(?<=gid=)\d+");
            return !match.Success;
        }
    }
}
