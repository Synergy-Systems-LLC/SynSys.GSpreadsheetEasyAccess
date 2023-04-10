using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

[assembly:InternalsVisibleTo("SynSys.GSpreadsheetEasyAccess.Tests")]
namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Represents one Google spreadsheet sheet.
    /// </summary>
    public class SheetModel
    {
        /// <summary>
        /// Sheet Name.
        /// </summary>
        [JsonProperty]
        public string Title { get; internal set; } = string.Empty;

        /// <summary>
        /// Google spreadsheet Id.
        /// </summary>
        [JsonProperty]
        public string SpreadsheetId { get; internal set; } = string.Empty;

        /// <summary>
        /// Google spreadsheet sheet Id.
        /// </summary>
        /// <remarks>
        /// The property has such a name because it correspondence
        /// to the gid parameter in the sheet uri.
        /// </remarks>
        [JsonProperty]
        public int Gid { get; internal set; } = -1;

        /// <summary>
        /// Spreadsheet name.
        /// </summary>
        [JsonProperty]
        public string SpreadsheetTitle { get; internal set; } = string.Empty;

        /// <summary>
        /// Key column name.
        /// </summary>
        [JsonProperty]
        public string KeyName { get; internal set; } = string.Empty;

        /// <summary>
        /// The mode by which work with the sheet is determined.
        /// </summary>
        [JsonProperty]
        public SheetMode Mode { get; internal set; }

        /// <summary>
        /// First row of the sheet.
        /// </summary>
        [JsonProperty]
        public List<string> Head { get; internal set; } = new List<string>();

        /// <summary>
        /// All rows included in this sheet except for the head.
        /// </summary>
        [JsonProperty]
        public List<Row> Rows { get; } = new List<Row>();

        /// <summary>
        /// Indicates that there are no rows in the sheet.
        /// </summary>
        /// <remarks>
        /// If the sheet has a head, then it is not taken into account.
        /// </remarks>
        /// <returns></returns>
        [JsonIgnore]
        public bool IsEmpty { get => Rows.Count == 0; }

        /// <summary>
        /// Adds an empty row to the end of the sheet.
        /// The row size will be equal to the maximum for this sheet
        /// </summary>
        public void AddRow()
        {
            AddRow(new List<string>());
        }

        /// <summary>
        /// Adds row to the end of the sheet.
        /// </summary>
        /// <remarks>
        /// The row size will be equal to the maximum for this sheet and if the data is larger than
        /// this size, then part of the data will not be included into the sheet.<br/>
        /// If there is less data, then the remaining cells will be filled with empty values.
        /// </remarks>
        /// <param name="data">Data to compose a row.</param>
        public void AddRow(IList<string> data)
        {
            AddRow(FindNextRowNumber(), Head.Count, data);
        }

        /// <summary>
        /// The method deletes the row only if it had a RowStatus.ToAppend status.<br/>
        /// Otherwise, the method does not delete the row, but assigns the RowStatus.ToDelete status.<br/>
        /// This status will be taken into account when deleting rows from a Google spreadsheet sheet.
        /// </summary>
        /// <param name="row">Row to delete</param>
        public void DeleteRow(Row row)
        {
            if (row.Status == RowStatus.ToAppend)
            {
                int nextRowNumber = row.Number + 1;

                Rows.Remove(row);
                DecreaseAllFollowingRowNumbersByOne(nextRowNumber);
            }
            else
            {
                row.Status = RowStatus.ToDelete;
            }
        }

        /// <summary>
        /// Assigning a deletion status to all rows
        /// and physical deletion of rows that have not yet been added.
        /// </summary>
        public void Clean()
        {
            for (int i = Rows.Count - 1; i >= 0; i--)
            {
                if (Rows[i].Status == RowStatus.ToAppend)
                {
                    Rows.RemoveAt(i);
                    continue;
                }

                Rows[i].Status = RowStatus.ToDelete;
            }
        }

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
            if (Mode == SheetMode.Simple)
            {
                return;
            }

            var lostHeaders = new List<string>();

            foreach (string columnHeader in requiredHeaders)
            {
                if (!Head.Contains(columnHeader))
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

        /// <summary>
        /// Merge with another version of the same sheet.
        /// </summary>
        /// <remarks>
        /// Sheet is considered the same if it has the same basic characteristics
        /// except for the list of rows.<br/>
        /// Row comparison is performed before merging. Row changes occur after comparison
        /// if needed.<br/>
        /// Cell values ​​and statuses are changed for rows, missing rows are added.<br/>
        /// </remarks>
        /// <param name="otherSheet">Same SheetModel</param>
        /// <exception cref="ArgumentNullException"/>
        /// <exception cref="ArgumentException"/>
        public void Merge(SheetModel otherSheet)
        {
            if (otherSheet == null)
            {
                throw new ArgumentNullException();
            }

            if (IsNotSameSheet(otherSheet, out string failReason))
            {
                throw new ArgumentException($"Sheets are not same. Reason: {failReason}");
            }

            int maximumRows = Math.Max(Rows.Count, otherSheet.Rows.Count);
            var rowsWithToAppendStatus = new List<Row>();

            for (int i = 0; i < maximumRows; i++)
            {
                Row currentRow = GetRowByIndex(i, Rows);
                Row otherRow = GetRowByIndex(i, otherSheet.Rows);

                if (MatchingRowNotFoundInOtherSheet(currentRow, otherRow))
                {
                    PrepareRowForDeletion(currentRow, rowsWithToAppendStatus);
                    continue;
                }

                if (MatchingRowNotFoundInCurrentSheet(currentRow, otherRow))
                {
                    AddRow(otherRow.GetData().Cast<string>().ToList());
                    continue;
                }

                MergeRows(currentRow, otherRow);
            }

            // These rows should be deleted because they cannot be given RowStatus.ToDelete status.
            // they cannot be given this status because they were not already in the Google spreadsheet sheet.
            DeleteRows(rowsWithToAppendStatus);
        }


        /// <summary>
        /// Initializes an empty sheet instance ready to be filled in.
        /// A sheet instance cannot be created outside the library.
        /// </summary>
        internal SheetModel() { }

        /// <summary>
        /// Filling the sheet with the creation of rows and cells.
        /// </summary>
        /// <param name="data">Data for sheet formation.</param>
        internal void Fill(IList<IList<object>> data)
        {
            IEnumerable<int> rowIndexes = Enumerable.Range(0, data.Count);
            int maxRowLength = GetMaxRowLength(data);

            foreach (int rowIndex in rowIndexes)
            {
                if (rowIndex == 0)
                {
                    if (Mode == SheetMode.Simple)
                    {
                        CreateEmptyHead(maxRowLength);
                    }
                    else
                    {
                        Head = data[0].Cast<string>().ToList();
                        continue;
                    }
                }

                var rowData = data[rowIndex].Cast<string>().ToList();
                AddRow(rowIndex + 1, maxRowLength, rowData, RowStatus.Original);
            }
        }

        /// <summary>
        /// Getting ValueRange for adding rows in Google spreadsheet sheet.
        /// </summary>
        /// <returns></returns>
        internal ValueRange GetAppendValueRange()
        {
            return new ValueRange
            {
                Values = GetAppendRows()
            };
        }

        /// <summary>
        /// Getting ValueRange from rows with ToChange status.
        /// </summary>
        /// <returns></returns>
        internal IList<ValueRange> GetChangeValueRange()
        {
            var valueRanges = new List<ValueRange>
            {
                new ValueRange()
                {
                    Values = new List<IList<object>>()
                }
            };

            var rowsToChange = Rows.FindAll(row => row.Status == RowStatus.ToChange);
            var previousRow = rowsToChange.First();
            rowsToChange.Remove(previousRow);

            valueRanges.Last().Values.Add(previousRow.GetData());
            valueRanges.Last().Range = $"{Title}!A{previousRow.Number}";

            foreach (var currentRow in rowsToChange)
            {
                if (currentRow.Number - previousRow.Number > 1)
                {
                    valueRanges.Add(new ValueRange()
                    {
                        Values = new List<IList<object>>(),
                        Range = $"{Title}!A{currentRow.Number}"
                    });
                }

                valueRanges.Last().Values.Add(currentRow.GetData());
                previousRow = currentRow;
            }

            return valueRanges;
        }

        /// <summary>
        /// Getting row groups to delete.
        /// </summary>
        internal List<List<Row>> GetDeleteRows()
        {
            var deleteGroups = new List<List<Row>>()
            {
                new List<Row>()
            };

            var rowsToDelete = Rows.FindAll(row => row.Status == RowStatus.ToDelete);
            // This list must be mirrored because the deletion of rows should occur from the end of the sheet.
            // Otherwise, indexes for subsequent deletions will fail.
            rowsToDelete.Reverse();
            var previousRow = rowsToDelete.First();
            rowsToDelete.Remove(previousRow);

            deleteGroups.Last().Add(previousRow);

            foreach (var currentRow in rowsToDelete)
            {
                if (previousRow.Number - currentRow.Number == 1) // Because numbering from largest to smallest
                {
                    deleteGroups.Last().Add(currentRow);
                }
                else
                {
                    deleteGroups.Add(new List<Row>() { currentRow });
                }

                previousRow = currentRow;
            }

            return deleteGroups;
        }

        /// <summary>
        /// Deleting rows with ToDelete status 
        /// so that after updating the data in the Google spreadsheet sheet, 
        /// you can use the same instance.
        /// </summary>
        internal void ClearDeletedRows()
        {
            var rowsToDelete = Rows.FindAll(row => row.Status == RowStatus.ToDelete);

            foreach (var row in rowsToDelete)
            {
                Rows.Remove(row);
            }
        }
 
        /// <summary>
        /// Renumbering all rows.<br/>
        /// Used when deleting some rows from a SheetModel without discerning which rows were deleted.
        /// </summary>
        internal void RenumberRows()
        {
            int number = FindFirstRowNumber();

            foreach (var row in Rows)
            {
                row.Number = number;
                number++;
            }
        }

        /// <summary>
        /// Reset the status of all rows to Original
        /// so that after updating the data in the Google spreadsheet sheet,
        /// you can use the same SheetModel instance.
        /// </summary>
        internal void ResetRowStatuses()
        {
            foreach (var row in Rows.FindAll(row => row.Status != RowStatus.Original))
            {
                row.Status = RowStatus.Original;
            }
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


        private void CreateEmptyHead(int length)
        {
            for (var i = 0; i < length; i++)
            {
                Head.Add(string.Empty);
            }
        }

        private void AddRow(int number, int length, IList<string> data, RowStatus status=RowStatus.ToAppend)
        {
            var row = new Row(data, length, Head)
            {
                Status = status,
                Number = number
            };

            if (!string.IsNullOrWhiteSpace(KeyName))
            {
                row.Key = row.Cells.Find(cell => cell.Title == KeyName);
            }

            Rows.Add(row);
        }

        /// <summary>
        /// Reducing all row numbers starting from the given one.<br/>
        /// Used when physically deleting one specific row.
        /// </summary>
        /// <param name="startRowNumber">Row number from which to start the reduction</param>
        private void DecreaseAllFollowingRowNumbersByOne(int startRowNumber)
        {
            foreach (var row in Rows.FindAll(r => r.Number >= startRowNumber))
            {
                row.Number -= 1;
            }
        }

        private int FindFirstRowNumber()
        {
            if (Mode == SheetMode.Head || Mode == SheetMode.HeadAndKey)
            {
                return 2;
            }

            return 1;
        }

        private int FindNextRowNumber()
        {
            if (Rows.Count > 0)
            {
                return Rows.Last().Number + 1;
            }
            else
            {
                return 1;
            }
        }

        private int GetMaxRowLength(IList<IList<object>> data)
        {
            if (Mode == SheetMode.Simple)
            {
                return data.Select(row => row.Count).Max();
            }
            else
            {
                return data.First().Count;
            }
        }

        private IList<IList<object>> GetAppendRows()
        {
            var data = new List<IList<object>>();

            foreach (var row in Rows)
            {
                if (row.Status == RowStatus.ToAppend)
                {
                    data.Add(row.GetData());
                }
            }

            return data;
        }

        private bool IsNotSameSheet(SheetModel otherSheet, out string failReason)
        {
            if (Title != otherSheet.Title)
            {
                failReason = nameof(otherSheet.Title);
                return true;
            }

            if (SpreadsheetId != otherSheet.SpreadsheetId)
            {
                failReason = nameof(otherSheet.SpreadsheetId);
                return true;
            }

            if (Gid != otherSheet.Gid)
            {
                failReason = nameof(otherSheet.Gid);
                return true;
            }

            if (SpreadsheetTitle != otherSheet.SpreadsheetTitle)
            {
                failReason = nameof(otherSheet.SpreadsheetTitle);
                return true;
            }

            if (KeyName != otherSheet.KeyName)
            {
                failReason = nameof(otherSheet.KeyName);
                return true;
            }

            if (Mode != otherSheet.Mode)
            {
                failReason = nameof(otherSheet.Mode);
                return true;
            }

            if (IsNotSameHead(otherSheet.Head))
            {
                failReason = nameof(otherSheet.Head);
                return true;
            }

            failReason = string.Empty;
            return false;
        }

        private bool IsNotSameHead(List<string> otherHead)
        {
            if (otherHead.Count != Head.Count)
            {
                return true;
            }

            for (int i = 0; i < Head.Count; i++)
            {
                if (Head[i] != otherHead[i])
                {
                    return true;
                }
            }

            return false;
        }

        private Row GetRowByIndex(int index, List<Row> rows)
        {
            if (index < rows.Count)
            {
                return rows[index];
            }

            return null;
        }

        private static bool MatchingRowNotFoundInCurrentSheet(Row currentRow, Row otherRow)
        {
            return currentRow == null && otherRow != null;
        }

        private static bool MatchingRowNotFoundInOtherSheet(Row currentRow, Row otherRow)
        {
            return currentRow != null && otherRow == null;
        }

        private static void PrepareRowForDeletion(Row currentRow, List<Row> rowsWithToAppendStatus)
        {
            if (currentRow.Status == RowStatus.ToAppend)
            {
                rowsWithToAppendStatus.Add(currentRow);
            }
            else
            {
                currentRow.Status = RowStatus.ToDelete;
            }
        }

        private static void MergeRows(Row currentRow, Row otherRow)
        {
            for (int j = 0; j < currentRow.Cells.Count; j++)
            {
                if (currentRow.Cells[j].Value != otherRow.Cells[j].Value)
                {
                    currentRow.Cells[j].Value = otherRow.Cells[j].Value;
                }
            }
        }

        private void DeleteRows(List<Row> rowsToAppend)
        {
            foreach (Row row in rowsToAppend)
            {
                Rows.Remove(row);
            }
        }
    }
}
