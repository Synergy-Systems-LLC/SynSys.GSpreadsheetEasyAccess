using System;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Represents a column of a Google Sheet.
    /// </summary>
    /// <remarks>
    /// Can only be named in A1 notation according to the Google Sheet.
    /// If a column with the same header already exists.
    /// </remarks>
    public class NativeColumn : AbstractColumn
    {
        /// <summary>
        /// Initialization of a spreadsheet column.
        /// </summary>
        /// <param name="number">Not index</param>
        public NativeColumn(int number)
        {
            _number = number;
            _title = GenerateA1NotationTitle(number);
        }

        /// <summary>
        /// Change number and title of the column.
        /// </summary>
        /// <remarks>
        /// This method includes changes of two components of the column,
        /// because title cannot be changed without its number.
        /// </remarks>
        /// <param name="number">Not index</param>
        public void ChangeNumber(int number)
        {
            _number = number;
            _title = GenerateA1NotationTitle(number);
        }

        /// <summary>
        /// Generate the column header in A1 notation based on its number.
        /// </summary>
        /// <param name="columnNumber">Not index</param>
        /// <returns>
        /// 1 -> А, 2 -> B, ... , 27 -> AA
        /// </returns>
        /// <exception cref="ArgumentException">
        /// If column number less than one.
        /// </exception>
        public static string GenerateA1NotationTitle(int columnNumber)
        {
            if (columnNumber < 1)
            {
                throw new ArgumentException(
                    message: $"Choosen column number \"{columnNumber}\" is less than one",
                    paramName: nameof(columnNumber)
                );
            }

            var alphabetLength = 26;
            var columnTitle = string.Empty;

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % alphabetLength;
                columnTitle = Convert.ToChar('A' + modulo).ToString() + columnTitle;
                columnNumber = (columnNumber - modulo) / alphabetLength;
            }

            return columnTitle;
        }
    }
}