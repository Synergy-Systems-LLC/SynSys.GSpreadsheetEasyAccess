using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    public class UserSheet : AbstractSheet
    {
        internal UserSheet() : base()
        {
            _firstRowNumber = 2;
            KeyName = string.Empty;
        }

        /// <summary>
        /// Key column name.
        /// </summary>
        [JsonProperty]
        public string KeyName { get; internal set; }

        /// <summary>
        /// Check if the required columns exist in the spreadsheet.
        /// </summary>
        /// <remarks>
        /// This method will throw an exception if the spreadsheet does not contain at least one required column.
        /// The method does not check for sheets with SheetMode.Simple.
        /// </remarks>
        /// <param name="requiredHeaders"></param>
        /// <exception cref="InvalidSheetHeadException"></exception>
        public void CheckHead(IEnumerable<string> requiredHeaders)
        {
            var lostHeaders = new List<string>();
            var headers = Head.Select(c => c.Title).ToList();

            foreach (string columnHeader in requiredHeaders)
            {
                if (!headers.Contains(columnHeader))
                {
                    lostHeaders.Add(columnHeader);
                }
            }

            if (lostHeaders.Count > 0)
            {
                throw new InvalidSheetHeadException(
                    $"in spreadsheet \"{SpreadsheetTitle}\" " +
                    $"in sheet\"{Title}\" " +
                    $"no required headers: " +
                    $"{string.Join(", ", lostHeaders)}."
                )
                {
                    Sheet = this,
                    LostHeaders = lostHeaders
                };
            }
        }

        public void AddColumn(string title)
        {
            var column = new UserColumn(Head.Count + 1, title)
            {
                Status = ChangeStatus.ToAppend
            };

            if (Head.Contains(column))
            {
                throw new ArgumentException(
                    message: $"Cannot add column \"{title}\" " +
                             $"due to column with same title exists",
                    paramName: nameof(title)
                );
            }

            Head.Add(column);
            AddCellInAllRows(column);
        }

        public void InsertColumn(int number, string title)
        {
            var column = new UserColumn(number, title)
            {
                Status = ChangeStatus.ToInsert
            };
            int columnIndex = column.Number - 1;

            if (Head.Contains(column))
            {
                throw new ArgumentException(
                    message: $"Столбец с именем {title} уже присудствует в шапке листа.",
                    paramName: nameof(title)
                );
            }

            Head.Insert(columnIndex, column);
            InsertCellInAllRows(column, columnIndex);
            RenumberColumnsAfterInserted(column.Number);
            ChangeCellsColumnReferenceAfterInserted(column.Number);
        }


        /// <exception cref="EmptySheetException"></exception>
        internal void ValidateData(IList<IList<object>> data)
        {
            if (data == null || data.Count == 0)
            {
                throw new EmptySheetException($"Sheet contains no data")
                {
                    Sheet = this
                };
            }
        }

        /// <exception cref="EmptySheetException"></exception>
        /// <exception cref="SheetKeyNotFoundException"></exception>
        internal void ValidateData(IList<IList<object>> data, string keyName)
        {
            ValidateData(data);

            if (!data[0].Contains(keyName))
            {
                throw new SheetKeyNotFoundException($"Sheet does not contain the key {keyName}")
                {
                    Sheet = this
                };
            }
        }

        internal override void CreateHead(List<List<string>> data)
        {
            CreateHead(data.First());
        }

        internal void CreateHead(IList<string> firstRow)
        {
            for (var columnNumber = 1; columnNumber <= firstRow.Count; columnNumber++)
            {
                int columnIndex = columnNumber - 1;
                string title = firstRow[columnIndex];
                bool columnIsKey = title == KeyName;

                Head.Add(new UserColumn(columnNumber, title, columnIsKey));
            }
        }


        protected override void AddRow(int number, int length, IList<string> data, ChangeStatus status=ChangeStatus.ToAppend)
        {
            var row = new Row(data, length, Head)
            {
                Status = status,
                Number = number
            };

            foreach (Cell cell in row.Cells)
            {
                var column = cell.Column as UserColumn;

                if (column.IsKey)
                {
                    row.Key = cell;
                    break;
                }
            }

            Rows.Add(row);
        }

        protected override int GetMaxRowLength(List<List<string>> data)
        {
            return data.First().Count;
        }

        protected bool IsNotSameSheet(UserSheet otherSheet, out string failReason)
        {
            bool result = base.IsNotSameSheet(otherSheet, out failReason);

            if (result == false)
            {
                return result;
            }

            if (KeyName != otherSheet.KeyName)
            {
                failReason = nameof(otherSheet.KeyName);
                return true;
            }

            return false;
        }
    }
}