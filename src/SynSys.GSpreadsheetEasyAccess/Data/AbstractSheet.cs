using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

[assembly:InternalsVisibleTo("SynSys.GSpreadsheetEasyAccess.Tests")]
namespace SynSys.GSpreadsheetEasyAccess.Data
{
    public abstract class AbstractSheet
    {
        protected int _firstRowNumber;

        /// <summary>
        /// Initializes an empty sheet instance ready to be filled in.
        /// A sheet instance cannot be created outside the library.
        /// </summary>
        internal AbstractSheet()
        {
            Title = string.Empty;
            Gid = -1;
            SpreadsheetId = string.Empty;
            SpreadsheetTitle = string.Empty;
            Head = new List<AbstractColumn>();
            Rows = new List<Row>();
        }

        /// <summary>
        /// Sheet Name.
        /// </summary>
        [JsonProperty]
        public string Title { get; internal set; }

        /// <summary>
        /// Google spreadsheet sheet Id.
        /// </summary>
        /// <remarks>
        /// The property has such a name because it correspondence
        /// to the gid parameter in the sheet uri.
        /// </remarks>
        [JsonProperty]
        public int Gid { get; internal set; }

        /// <summary>
        /// Google spreadsheet Id.
        /// </summary>
        [JsonProperty]
        public string SpreadsheetId { get; internal set; }

        /// <summary>
        /// Spreadsheet name.
        /// </summary>
        [JsonProperty]
        public string SpreadsheetTitle { get; internal set; }

        /// <summary>
        /// First row of the sheet.
        /// </summary>
        [JsonProperty]
        public List<AbstractColumn> Head { get; internal set; }

        /// <summary>
        /// All rows included in this sheet except for the head.
        /// </summary>
        [JsonProperty]
        public List<Row> Rows { get; internal set; }

        /// <summary>
        /// Indicates that there are no rows in the sheet.
        /// </summary>
        /// <remarks>
        /// If the sheet has a head, then it is not taken into account.
        /// </remarks>
        /// <returns></returns>
        [JsonIgnore]
        public bool IsEmpty => Rows.Count == 0;

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
        /// <param name="number">Number of row to delete</param>
        public void DeleteRow(int number)
        {
            var row = Rows.Find(r => r.Number == number);

            if (row is null)
            {
                throw new ArgumentException($"the row with the number {number} was not found", nameof(number));
            }

            DeleteRow(row);
        }

        /// <summary>
        /// The method deletes the row only if it had a RowStatus.ToAppend status.<br/>
        /// Otherwise, the method does not delete the row, but assigns the RowStatus.ToDelete status.<br/>
        /// This status will be taken into account when deleting rows from a Google spreadsheet sheet.
        /// </summary>
        /// <param name="row">Row to delete</param>
        public void DeleteRow(Row row)
        {
            if (row.Status == ChangeStatus.ToAppend)
            {
                int nextRowNumber = row.Number + 1;

                Rows.Remove(row);
                DecreaseAllFollowingRowNumbersByOne(nextRowNumber);
            }
            else
            {
                row.Status = ChangeStatus.ToDelete;
            }
        }

        /// <summary>
        /// Отменить удаление строки по её номеру.
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если строки с таким номером не существует.
        /// Если строки с таким номером не планировалось удалять.
        /// </exception>
        public void UndoRowDeletion(int number)
        {
            Row row = TryToFindRow(number);
            UndoRowDeletion(row);
        }

        /// <summary>
        /// Отменить удаление.
        /// </summary>
        /// <param name="row"></param>
        /// <exception cref="InvalidOperationException">
        /// Если строка не планировалась удаляться.
        /// </exception>
        public void UndoRowDeletion(Row row)
        {
            if (row.Status != ChangeStatus.ToDelete)
            {
                // TODO дописать
                throw new InvalidOperationException();
            }

            row.Status = ChangeStatus.Original;
        }

        /// <summary>
        /// Удалить столбец по его номеру.
        /// </summary>
        /// <remarks>
        /// Столбец на самом деле не удаляется, ему назначается статус на удаление.
        /// </remarks>
        /// <param name="number">Не индекс</param>
        /// <exception cref="InvalidOperationException">
        /// Если столбец с таким номером не планировался удаляться.
        /// </exception>
        public void DeleteColumn(int number)
        {
            var column = Head.Find(c => c.Number == number);

            if (column is null)
            {
                throw new ArgumentException($"the column with the number {number} was not found", nameof(number));
            }

            DeleteColumn(column);
        }

        // TODO дописать
        /// <summary>
        /// 
        /// </summary>
        /// <param name="column"></param>
        public void DeleteColumn(AbstractColumn column)
        {
            if (column.Status == ChangeStatus.ToAppend)
            {
                int nextRowNumber = column.Number + 1;

                Head.Remove(column);
                // TODO написать этот метод
                // DecreaseAllFollowingColumnNumbersByOne(nextRowNumber);
            }
            else
            {
                column.Status = ChangeStatus.ToDelete;
            }
        }

        /// <summary>
        /// Отменить удаление столбца по его номеру.
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если столбца с таким номером не существует.
        /// Если столбец с таким номером не планировался удаляться.
        /// </exception>
        public void UndoColumnDeletion(int number)
        {
            AbstractColumn column = TryToFindColumn(number);
            UndoColumnDeletion(column);
        }

        /// <summary>
        /// Отменить удаление.
        /// </summary>
        /// <param name="column"></param>
        /// <exception cref="InvalidOperationException">
        /// Если столбец не планировался удаляться.
        /// </exception>
        public void UndoColumnDeletion(AbstractColumn column)
        {
            if (column.Status != ChangeStatus.ToDelete)
            {
                // TODO дописать
                throw new InvalidOperationException();
            }

            column.Status = ChangeStatus.Original;
        }

        /// <summary>
        /// Assigning a deletion status to all rows
        /// and physical deletion of rows that have not yet been added.
        /// </summary>
        public void Clean()
        {
            for (int i = Rows.Count - 1; i >= 0; i--)
            {
                if (Rows[i].Status == ChangeStatus.ToAppend)
                {
                    Rows.RemoveAt(i);
                    continue;
                }

                Rows[i].Status = ChangeStatus.ToDelete;
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
        public void Merge(AbstractSheet otherSheet)
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
        /// Filling the sheet with the creation of rows and cells.
        /// </summary>
        /// <param name="data">Data for sheet formation.</param>
        internal void Fill(IList<IList<object>> data)
        {
            var rowsData = data.Select(r => r.Select(c => (string)c).ToList()).ToList();
            int maxRowLength = GetMaxRowLength(rowsData);

            CreateHead(rowsData);

            int countRows = rowsData.Count - (_firstRowNumber - 1);

            foreach (int number in Enumerable.Range(_firstRowNumber, countRows))
            {
                int index = number - 1;
                var rowData = rowsData[index];
                AddRow(number, maxRowLength, rowData, ChangeStatus.Original);
            }
        }

        /// <summary>
        /// Getting ValueRange for adding rows in Google spreadsheet sheet.
        /// </summary>
        /// <returns></returns>
        // TODO переименовать в GetAppendRowsValueRange
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
                // я забыл зачем инициализировать пустой ValueRange
                new ValueRange()
                {
                    Values = new List<IList<object>>()
                }
            };

            // я не понимаю зачем все эти действия
            // почему нельзя сразу начать с цикла?
            // надо подумать о рефакторинге.
            var rowsToChange = Rows.FindAll(row => row.Status == ChangeStatus.ToChange);
            var previousRow = rowsToChange.First();
            rowsToChange.Remove(previousRow);

            valueRanges.Last().Values.Add(previousRow.GetData());
            valueRanges.Last().Range = $"{Title}!A{previousRow.Number}";

            foreach (var currentRow in rowsToChange)
            {
                if (currentRow.Number - previousRow.Number > 1)
                {
                    valueRanges.Add(
                        new ValueRange()
                        {
                            Values = new List<IList<object>>(),
                            Range = $"{Title}!A{currentRow.Number}"
                        }
                    );
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

            var rowsToDelete = Rows.FindAll(row => row.Status == ChangeStatus.ToDelete);
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
            var rowsToDelete = Rows.FindAll(row => row.Status == ChangeStatus.ToDelete);

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
            int number = _firstRowNumber;

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
            // TODO переписать полностью на Linq
            foreach (var row in Rows.FindAll(row => row.Status != ChangeStatus.Original))
            {
                row.Status = ChangeStatus.Original;
            }
        }

        internal void ResetColumnStatuses()
        {
            Head.ForEach(columns => columns.Status = ChangeStatus.Original); 
        }

        internal List<List<int>> GetDeleteColumns()
        {
            var deleteGroups = new List<List<int>>()
            {
                new List<int>()
            };

            return deleteGroups;
        }

        internal int GetCountAppendColumns()
        {
            return Head.Where(c => c.Status == ChangeStatus.ToAppend).Count();
        }

        internal AbstractColumn GetFirstAddedColumn()
        {
            return Head.Where(c => c.Status == ChangeStatus.ToAppend).FirstOrDefault();
        }

        internal List<IList<object>> GetAddedColumnsValues()
        {
            var values = new List<IList<object>>();

            List<AbstractColumn> addedColumns = Head
                .Where(c => c.Status == ChangeStatus.ToAppend)
                .ToList();

            values.Add(addedColumns.Select(column => column.Title).Cast<object>().ToList());

            foreach (Row row in Rows)
            {
                var rowValues = new List<object>();

                foreach (AbstractColumn column in addedColumns)
                {
                    rowValues.Add(row.Cells[column.Number - 1].Value);
                }

                values.Add(rowValues);
            }

            return values;
        }

        internal List<List<AbstractColumn>> GetInsertColumnsGroups()
        {
            // Логика построена таким образом,
            // чтобы если в листе нет вставляемых столбцов, метод вернул пустрой список групп.
            // Это нужно, чтобы внешний код не выяснял есть ли в первой группе, столбцы.
            // Если нет групп, то и столбцов нет, а если группа есть, то в ней точно будут столбцы.
            var groups = new List<List<AbstractColumn>>();
            var group = new List<AbstractColumn>();

            List<AbstractColumn> insertColumns = Head
                .Where(c => c.Status == ChangeStatus.ToInsert)
                .ToList();

            foreach (AbstractColumn column in insertColumns)
            {
                if (group.Count == 0 || IsCurrentColumnGreaterPreviousByOne(group.Last(), column))
                {
                    group.Add(column);
                    continue;
                }

                groups.Add(group);
                group = new List<AbstractColumn>() { column };
            }

            if (group.Any())
            {
                groups.Add(group);
            }

            return groups;
        }

        internal abstract void CreateHead(List<List<string>> data);


        protected abstract void AddRow(int number, int length, IList<string> data, ChangeStatus status = ChangeStatus.ToAppend);

        protected abstract int GetMaxRowLength(List<List<string>> data);

        protected int FindNextRowNumber()
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

        protected void AddCellInAllRows(AbstractColumn column)
        {
            foreach (Row row in Rows)
            {
                row.Cells.Add(new Cell(string.Empty, column, row));
            }
        }

        protected void InsertCellInAllRows(AbstractColumn column, int columnIndex)
        {
            foreach (Row row in Rows)
            {
                row.Cells.Insert(columnIndex, new Cell(string.Empty, column, row));
            }
        }

        /// <exception cref="InvalidOperationException"></exception>
        protected AbstractColumn TryToFindColumn(string title)
        {
            return Head.Find(c => c.Title.Contains(title)) ?? throw new InvalidOperationException();
        }

        /// <exception cref="InvalidOperationException"></exception>
        protected Row TryToFindRow(int number)
        {
            return Rows.Find(r => r.Number == number) ?? throw new InvalidOperationException();
        }

        /// <exception cref="InvalidOperationException"></exception>
        protected AbstractColumn TryToFindColumn(int number)
        {
            if (number < 1 || number > Head.Count)
            {
                throw new InvalidOperationException(
                    $"Попытка удалить столбец с некорректным номером {number}"
                );
            }

            return Head[number];
        }

        protected void RenumberColumnsAfterInserted(int insertedColumnNumber)
        {
            for (int columnIndex = insertedColumnNumber; columnIndex < Head.Count ; columnIndex++)
            {
                Head[columnIndex].ChangeNumber(columnIndex + 1);
            }
        }

        protected void ChangeCellsColumnReferenceAfterInserted(int insertedColumnNumber)
        {
            for (int columnIndex = insertedColumnNumber; columnIndex < Head.Count; columnIndex++)
            {
                AbstractColumn column = Head[columnIndex];

                foreach (Row row in Rows)
                {
                    row.Cells[columnIndex].Column = column;
                }
            }
        }


        private bool IsCurrentColumnGreaterPreviousByOne(AbstractColumn previous, AbstractColumn current)
        {
            return previous.Number + 1 == current.Number;
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

        private IList<IList<object>> GetAppendRows()
        {
            var data = new List<IList<object>>();

            foreach (var row in Rows)
            {
                if (row.Status == ChangeStatus.ToAppend)
                {
                    data.Add(row.GetData());
                }
            }

            return data;
        }

        // TODO надо подумать для каких листов это валидная работа
        protected bool IsNotSameSheet(AbstractSheet otherSheet, out string failReason)
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

            if (IsNotSameHead(otherSheet.Head))
            {
                failReason = nameof(otherSheet.Head);
                return true;
            }

            failReason = string.Empty;
            return false;
        }

        private bool IsNotSameHead(List<AbstractColumn> otherHead)
        {
            if (otherHead.Count != Head.Count)
            {
                return true;
            }

            for (var i = 0; i < Head.Count; i++)
            {
                if (Head[i].Title != otherHead[i].Title)
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
            if (currentRow.Status == ChangeStatus.ToAppend)
            {
                rowsWithToAppendStatus.Add(currentRow);
            }
            else
            {
                currentRow.Status = ChangeStatus.ToDelete;
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
