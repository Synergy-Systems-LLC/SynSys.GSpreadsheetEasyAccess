using System;
using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Represents a column that starts at one of the rows in a Google Sheet.
    /// </summary>
    /// <remarks>
    /// A user-defined column has a header assigned by the user,
    /// and can also be a key for this table.
    /// </remarks>
    public class UserColumn : AbstractColumn
    {
        /// <summary>
        /// Initialization of a spreadsheet user column.
        /// </summary>
        /// <param name="number"></param>
        /// <param name="title"></param>
        /// <param name="isKey"></param>
        public UserColumn(int number, string title, bool isKey = false)
        {
            Number = number;
            Title = title;
            IsKey = isKey;
        }

        /// <summary>
        /// Is the column considered a key.
        /// </summary>
        public bool IsKey { get; }

        public override void ChangeNumber(int newNumber)
        {
            Number = newNumber;
        }

        public override bool Equals(AbstractColumn other)
        {
            // В сравнении нет Number потому что главным определяющим параметром является заголовок.
            // Столбец может находиться на любой позиции, но заголовок из-за этого меняться не будет.
            return other is UserColumn && Title == other.Title;
        }

        /// <summary>
        /// Change the column title depending of existing ones.
        /// </summary>
        /// <remarks>
        /// There should be no columns with the same title in the sheet.
        /// </remarks>
        /// <param name="newTitle"></param>
        /// <param name="existsColumns"></param>
        /// <exception cref="ArgumentException">
        /// If a column with the same title already exists.
        /// </exception>
        public void ChangeTitle(string newTitle, IEnumerable<UserColumn> existsColumns)
        {
            if (existsColumns.Select(c => c.Title).Contains(newTitle))
            {
                throw new ArgumentException(
                    message: $"Cannot rename column \"{Title}\" " +
                             $"due to column with same title \"{newTitle}\" exists",
                    paramName: nameof(newTitle)
                );
            }

            Title = newTitle;
        }
    }
}