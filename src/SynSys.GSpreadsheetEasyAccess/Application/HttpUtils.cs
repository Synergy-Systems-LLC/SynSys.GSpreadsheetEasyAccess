using System.Text.RegularExpressions;

namespace SynSys.GSpreadsheetEasyAccess.Application
{
    /// <summary>
    /// Provide static methods for parsing Google spreadsheets uri.
    /// </summary>
    public static class HttpUtils
    {
        /// <summary>
        /// Get the Google spreadsheet Id.
        /// </summary>
        /// <param name="uri">Full Google spreadsheet sheet uri</param>
        /// <returns>
        /// String of 44 characters following d/
        /// </returns>
        public static string GetSpreadsheetIdFromUri(string uri)
        {
            var match = Regex.Match(uri, @"(?s)(?<=d/).+?(?=/edit)");
            return match.ToString();
        }

        /// <summary>
        /// Get the Google spreadsheet sheet Id.
        /// </summary>
        /// <param name="uri">Full Google spreadsheet sheet uri</param>
        /// <returns>
        /// Returns -1 if there is no gid in uri.
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
        /// Checking that given uri not belongs to Google spreadsheet sheet.
        /// </summary>
        /// <param name="uri">Full uri Google spreadsheet sheet.</param>
        public static bool IsNotCorrectUri(string uri)
        {
            return string.IsNullOrEmpty(uri)
                || NotExistsSpreadsheet(uri)
                || NotExistsId(uri)
                || NotExistsGid(uri);
        }


        private static bool NotExistsSpreadsheet(string url)
        {
            return !url.Contains("spreadsheets");
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
