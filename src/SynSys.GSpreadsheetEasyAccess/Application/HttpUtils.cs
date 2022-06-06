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
    }
}
