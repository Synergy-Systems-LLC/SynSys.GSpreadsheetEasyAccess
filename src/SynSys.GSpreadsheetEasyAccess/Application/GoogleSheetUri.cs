using System;
using System.Text.RegularExpressions;

namespace SynSys.GSpreadsheetEasyAccess.Application
{
    public class GoogleSheetUri : Uri
    {
        public GoogleSheetUri(string uri) : base(uri)
        {
            CheckGoogleDocsDomain(uri);
            CheckSpreadsheetAffiliation(uri);
            SpreadsheetId = GetSpreadsheetId(uri);
            SheetId = GetGid(uri);
        }

        public string SpreadsheetId { get; }

        public int SheetId { get; }


        private void CheckGoogleDocsDomain(string uri)
        {
            var domain = "docs.google.com";

            if(uri.Contains(domain) == false)
            {
                throw new UriFormatException($"This is not a google docs domain.");
            }
        }

        private void CheckSpreadsheetAffiliation(string uri)
        {
            var affiliation = "spreadsheets";

            if(uri.Contains(affiliation) == false)
            {
                throw new UriFormatException($"There is no \'{affiliation}\' in the uri.");
            }
        }

        private string GetSpreadsheetId(string uri)
        {
            // String of characters following d/
            var match = Regex.Match(uri, @"(?s)(?<=d/).+?(?=/edit)");

            if (match.Success == false)
            {
                throw new UriFormatException("Spreadsheet id is missing.");
            }

            return match.ToString();
        }

        private int GetGid(string uri)
        {
            var match = Regex.Match(uri, @"(?s)(?<=gid=)\d+");

            if (match.Success == false)
            {
                throw new UriFormatException("Sheet gid is missing.");
            }

            return int.Parse(match.Value);
        }
    }
}
