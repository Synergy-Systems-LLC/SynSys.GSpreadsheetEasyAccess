using System.Text.RegularExpressions;

namespace SynSys.GSpreadsheetEasyAccess.Application
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
    }
}
