using System;
using System.Net;
using System.Text.RegularExpressions;

namespace GetGoogleSheetDataAPI
{
    public static class HttpManager
    {
        public static string Status { get; internal set; }

        public static bool IsCorrectUrl(string url)
        {
            if (String.IsNullOrEmpty(url))
            {
                Status = "Url пустой";
                return false;
            }

            if (NotExistsSpreadsheet(url))
            {
                Status = "Url адрес не принадлежит google spreadsheet";
                return false;
            }

            if (NotExistsId(url))
            {
                Status = "В Url адресе некорректный spreadsheetId";
                return false;
            }

            if (NotExistsGid(url))
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

        public static string GetSpreadsheetIdFromUrl(string url)
        {
            var match = Regex.Match(url, @"(?s)(?<=d/).+?(?=/edit)");
            return match.ToString();
        }

        public static string GetGidFromUrl(string url)
        {
            var match = Regex.Match(url, @"(?s)(?<=gid=)\d+");
            return match.ToString();
        }

        public static string GetHttpStatusCode(string errorMessage)
        {
            string pattern = @"\[(\d+)\]";
            var rx = new Regex(pattern); //один из вариантов

            if (rx.IsMatch(errorMessage))
            {
                return rx.Match(errorMessage).Groups[1].Value;
            }
            else
            {
                return string.Empty;
            }
        }

        public static string GetMessageByCode(string httpCode)
        {
            switch (httpCode)
            {
                case "400":
                    return "Неверный запрос. Url адрес не существует";
                case "401":
                    return "Проблемы с авторизацией в google";
                case "403":
                    return "Отказ в доступе. Обратитесь к владельцу таблицы google";
                default:
                    return $"http status code - {httpCode}";
            }
        }
    }
}
