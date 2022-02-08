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
    /// Перечисление для определённого заполнения листа.
    /// Выбор режима влияет на состояние ячеек и строк листа.
    /// </summary>
    public enum SheetMode
    {
        /// <summary>
        /// Таблица без шапки и ключа
        /// </summary>
        Simple,
        /// <summary>
        /// Таблица с шапкой в одну строку
        /// </summary>
        Head,
        /// <summary>
        /// Таблица с шапкой в одну строку и столбцом-ключём
        /// </summary>
        HeadAndKey
    }

    /// <summary>
    /// Тип представляет один лист гугл таблицы.
    /// </summary>
    public class SheetModel
    {
        /// <summary>
        /// Имя листа
        /// </summary>
        [JsonProperty]
        public string Title { get; internal set; } = string.Empty;

        /// <summary>
        /// Id google листа
        /// </summary>
        [JsonProperty]
        public string SpreadsheetId { get; internal set; } = string.Empty;

        /// <summary>
        /// Id листа
        /// </summary>
        /// <remarks>
        /// Свойство имеет такое название потому что в uri листа ему
        /// соответствует параметр gid.
        /// </remarks>
        [JsonProperty]
        public string Gid { get; internal set; } = string.Empty;

        /// <summary>
        /// Имя листа
        /// </summary>
        [JsonProperty]
        public string SpreadsheetTitle { get; internal set; } = string.Empty;

        /// <summary>
        /// Текущий статус.<br/>
        /// <c>Пример: Лист не содержит данных</c>
        /// </summary>
        [JsonProperty]
        public string Status { get; internal set; } = string.Empty;

        /// <summary>
        /// Имя ключевого столбца, если он есть.
        /// </summary>
        [JsonProperty]
        public string KeyName { get; internal set; } = string.Empty;

        /// <summary>
        /// Режим, по которому определяется работа с листом.
        /// </summary>
        [JsonProperty]
        public SheetMode Mode { get; internal set; }

        /// <summary>
        /// Шапка (первая строка листа) если есть
        /// </summary>
        [JsonProperty]
        public List<string> Head { get; internal set; } = new List<string>();

        /// <summary>
        /// Все строки входящие в данный лист
        /// за исключением шапки
        /// </summary>
        [JsonProperty]
        public List<Row> Rows { get; } = new List<Row>();

        /// <summary>
        /// Показывает пустой ли лист.
        /// </summary>
        /// <returns></returns>
        [JsonIgnore]
        public bool IsEmpty { get => Rows.Count == 0; }

        /// <summary>
        /// Добавляет пустую строку в конец листа.
        /// Размер строки будет равен максимальному для данного листа.
        /// </summary>
        public void AddRow()
        {
            AddRow(new List<string>());
        }

        /// <summary>
        /// Добавляет строку в конец листа.
        /// Размер строки будет равен максимальному для данной листа
        /// и если данных будет больше чем этот размер, то часть данных не попадёт в лист.
        /// если данных будет меньше, то строка дозаполнится пустыми значениями.
        /// </summary>
        /// <param name="data">Данные для составленния строки</param>
        public void AddRow(List<string> data)
        {
            AddRow(FindNextRowNumber(), Head.Count, data);
        }

        /// <summary>
        /// Метод удаляет строку только если у неё был статус RowStatus.ToAppend.<br/>
        /// В противном случае метод не удаляет строку, а назначает ей статус на удаление RowStatus.ToDelete.<br/>
        /// Данный статус будет учитываться при удалении строк из листа в google.
        /// </summary>
        /// <param name="row">Удаляемая строка</param>
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
        /// Назначение всем строкам статус на удаление и 
        /// физическое удаление ещё не добавленных строк.
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
        /// Проверка шапки.
        /// </summary>
        /// <param name="requiredHeaders"></param>
        /// <param name="lostedHeaders"></param>
        /// <returns>
        /// true если нет хотя бы одного заголовка; в противном случае false.
        /// </returns>
        public bool DoesNotContainsSomeHeaders(
            IEnumerable<string> requiredHeaders, out List<string> lostedHeaders)
        {
            bool result = false;
            lostedHeaders = new List<string>();

            foreach (string header in requiredHeaders)
            {
                if (!Head.Contains(header))
                {
                    lostedHeaders.Add(header);
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Слияние с другой версией этого же листа.<br/>
        /// Лист считается таким же если у него совпадают все основные характеристики 
        /// кроме списка строк.<br/>
        /// Перед слиянием выполняется сравнение строк. После сравнения происходят изменения строк
        /// если нужно.<br/>
        /// У строк меняются значения ячеек и статусы, добавляются недостающие строки.<br/>
        /// </summary>
        /// <param name="otherSheet"></param>
        /// <exception cref="ArgumentNullException"/>
        /// <exception cref="ArgumentException">
        /// </exception>
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

            // Эти строки надо удалять потому что им нельзя дать статус RowStatus.ToDelete.
            // А данный статус им нельзя давать потому что их и так не было в гугл таблице.
            DeleteRows(rowsWithToAppendStatus);
        }


        /// <summary>
        /// Инициализирует пустой экземпеляр листа готовый для заполнения.
        /// Экземпляр листа нельзя создавать вне библиотеки.
        /// </summary>
        internal SheetModel() { }

        /// <summary>
        /// Заполнение листа с созданием строк и ячеек.
        /// </summary>
        /// <param name="data">Данные для формирования листа</param>
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
        /// Получение ValueRange для добавления строк в google spreadsheet.
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
        /// Получение ValueRange из строк со статусом ToChange.
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
        /// Получение групп строк для удаления.
        /// </summary>
        /// <returns></returns>
        internal List<List<Row>> GetDeleteRows()
        {
            var deleteGroups = new List<List<Row>>()
            {
                new List<Row>()
            };

            var rowsToDelete = Rows.FindAll(row => row.Status == RowStatus.ToDelete);
            // Данный список нужно зеркалить потому что удаление строк должно происходить с конца листа.
            // В противном случае собьются индексы для последующих удалений.
            rowsToDelete.Reverse();
            var previousRow = rowsToDelete.First();
            rowsToDelete.Remove(previousRow);

            deleteGroups.Last().Add(previousRow);

            foreach (var currentRow in rowsToDelete)
            {
                if (previousRow.Number - currentRow.Number == 1) // Потому что нумерация от большего к меньшему
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
        /// Удаление строк со статусом ToDelete
        /// чтобы после обновления данных в гугл таблице можно было пользоваться
        /// тем же инстансем типа Sheet.
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
        /// Перенумерация всех строк.<br/>
        /// Используется при удалении некоторых строк из листа, при этом не ясно какие именно строки
        /// были удалены.
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
        /// Сброс статуса всех строк на Original
        /// чтобы после обновления данных в гугл таблице можно было пользоваться
        /// тем же инстансем SheetModel.
        /// </summary>
        internal void ResetRowStatuses()
        {
            foreach (var row in Rows.FindAll(row => row.Status != RowStatus.Original))
            {
                row.Status = RowStatus.Original;
            }
        }

        internal static SheetModel Create(Spreadsheet spreadsheet, Sheet sheet, SheetMode mode, string keyName, IList<IList<object>> data)
        {
            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheet.SpreadsheetId,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Gid = sheet.Properties.SheetId.ToString(),
                Title = sheet.Properties.Title,
                Mode = mode,
                KeyName = keyName
            };

            sheetModel.Fill(data);

            return sheetModel;
        }

        /// <exception cref="InvalidSheetUriException">
        /// Если uri не корректный
        /// </exception>
        internal static void CheckUri(string uri)
        {
            if (string.IsNullOrWhiteSpace(uri))
            {
                throw new InvalidSheetUriException("uri не может быть пустым");
            }
        }

        /// <exception cref="InvalidSheetGidException">
        /// Если uri не корректный
        /// </exception>
        internal static void CheckGid(int gid)
        {
            if (gid < 0)
            {
                throw new InvalidSheetGidException($"gid листа не может быть отрицательным числом");
            }
        }

        /// <exception cref="InvalidSheetNameException">
        /// Если uri не корректный
        /// </exception>
        internal static void CheckName(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new InvalidSheetNameException($"Не существует листа с пустым именем");
            }
        }

        /// <exception cref="EmptySheetException">
        /// Если uri не корректный
        /// </exception>
        internal static void ValidateData(IList<IList<object>> data)
        {
            if (data == null || data.Count == 0)
            {
                throw new EmptySheetException($"Лист не содержит данных");
            }
        }

        /// <exception cref="KeyNotFoundException">
        /// Если uri не корректный
        /// </exception>
        internal static void ValidateData(IList<IList<object>> data, string keyName)
        {
            ValidateData(data);

            if (!data[0].Contains(keyName))
            {
                throw new KeyNotFoundException($"Лист не содержит ключ {keyName}");
            }
        }


        private void CreateEmptyHead(int length)
        {
            for (var i = 0; i < length; i++)
            {
                Head.Add(string.Empty);
            }
        }

        private void AddRow(
            int number, int length, List<string> data, RowStatus status=RowStatus.ToAppend)
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
        /// Уменьшение всех строк начиная с заданной.<br/>
        /// Используется при физическом удалении одной конкретной строки.
        /// </summary>
        /// <param name="startRowNumber">Номер строки с которой надо начать уменьшение</param>
        private void DecreaseAllFollowingRowNumbersByOne(int startRowNumber)
        {
            foreach (var row in Rows.FindAll(r => r.Number >= startRowNumber))
            {
                row.Number -= 1;
            }
        }

        private int FindFirstRowNumber()
        {
            if (Mode == SheetMode.Head
                || Mode == SheetMode.HeadAndKey)
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

            if (Status != otherSheet.Status)
            {
                failReason = nameof(otherSheet.Status);
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
