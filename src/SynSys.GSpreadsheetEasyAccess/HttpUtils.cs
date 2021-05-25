using System.Text.RegularExpressions;

namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Класс функций для работы с uri и http статус кодами.
    /// </summary>
    public static class HttpUtils
    {
        /// <summary>
        /// Получение идентификатора гугл таблицы.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа гугл таблицы</param>
        /// <returns></returns>
        public static string GetSpreadsheetIdFromUri(string uri)
        {
            var match = Regex.Match(uri, @"(?s)(?<=d/).+?(?=/edit)");
            return match.ToString();
        }

        /// <summary>
        /// Получение идентификатора листа гугл таблицы.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа гугл таблицы</param>
        /// <returns>
        /// Возвращает -1 если в uri нет gid.
        /// </returns>
        public static int GetGidFromUri(string uri)
        {
            var match = Regex.Match(uri, @"(?s)(?<=gid=)\d+");

            if (string.IsNullOrWhiteSpace(match.Value))
            {
                return -1;
            }

            return int.Parse(match.Value);
        }

        /// <summary>
        /// Проверка корректности полного uri листа гугл таблицы.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа гугл таблицы</param>
        /// <param name="status">Информация о том почему uri не корректный</param>
        /// <returns></returns>
        public static bool IsCorrectUri(string uri, out string status)
        {
            if (string.IsNullOrEmpty(uri))
            {
                status = "Url пустой";
                return false;
            }

            if (NotExistsSpreadsheet(uri))
            {
                status = "Url адрес не принадлежит google spreadsheet";
                return false;
            }

            if (NotExistsId(uri))
            {
                status = "В Url адресе некорректный spreadsheetId";
                return false;
            }

            if (NotExistsGid(uri))
            {
                status = "В Url адресе некорректный gid листа или он отсутствует";
                return false;
            }

            status = string.Empty;
            return true;
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
