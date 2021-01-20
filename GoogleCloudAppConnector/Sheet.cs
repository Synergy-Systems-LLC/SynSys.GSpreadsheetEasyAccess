using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Linq;

namespace GetGoogleSheetDataAPI
{
    /// <summary>
    /// Виды таблиц:
    /// 1. Без шапки и без ключа
    /// 2. Шапка (1 строка)
    /// 3. Шапка (1 строка) + Ключ (1 столцеб)
    /// 4. Шапка (>1 строки) + Составной ключ (>1 столбца)
    /// </summary>
    public class Sheet
    {
        public string Title { get; internal set; } = string.Empty;
        public string SpreadsheetId { get; internal set; } = string.Empty;
        public string Gid { get; internal set; } = string.Empty;
        public string SpreadsheetTitle { get; internal set; }
        public string Status { get; internal set; } = string.Empty;
        public List<Row> Rows { get; internal set; } = new List<Row>();

        /// <summary>
        /// Инициализирует пустой экземпеляр таблицы готовый для заполнения.
        /// Экземпляр таблицы нельзя создавать вне библиотеки.
        /// </summary>
        internal Sheet() { }

        /// <summary>
        /// Заполнение таблицы с созданием строк и ячеек.
        /// </summary>
        /// <param name="data">Данные для формирования таблицы</param>
        internal void Fill(IList<IList<object>> data)
        {
            var rows = Enumerable.Range(0, data.Count);
            var maxColumns = data.First().Count;

            foreach (int rowIndex in rows)
            {
                var rowData = data[rowIndex].Cast<string>().ToList();
                AddRow(rowIndex, maxColumns, rowData, RowStatus.Original);
            }
        }

        /// <summary>
        /// Добавляет пустую строку в конец таблицы.
        /// Размер строки будет равен максимальному для данной таблицы.
        /// </summary>
        public void AddRow()
        {
            AddRow(new List<string>());
        }

        /// <summary>
        /// Добавляет пустую строку в конец таблицы.
        /// Размер строки будет равен максимальному для данной таблицы
        /// и если данных будет больше чем этот размер, то часть данных не попадёт в таблицу.
        /// если данных будет меньше, то строка дозаполнится пустыми значениями.
        /// </summary>
        /// <param name="rowData">Данные для составленния строки</param>
        public void AddRow(List<string> rowData)
        {
            AddRow(
                Rows.Count,
                Rows.First().Cells.Count,
                rowData
            );
        }

        /// <summary>
        /// Основной метод для добавления строки в конец таблицы.
        /// </summary>
        /// <param name="index">Индекс данной строки</param>
        /// <param name="maxLength">Максимальная длинаа строки</param>
        /// <param name="rowData">Данные для формированния строки</param>
        /// <param name="status">
        /// Статус строки. По умолчанию это ToAppend, но есть возможность его изменнить
        /// </param>
        private void AddRow(
            int index,
            int maxLength,
            List<string> rowData,
            RowStatus status=RowStatus.ToAppend)
        {
            var row = new Row(rowData, maxLength)
            {
                Status = status,
                Number = index + 1
            };

            Rows.Add(row);
        }

        /// <summary>
        /// Метод для получение ValueRange для добавления строк в google таблицу.
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
        /// Метод преобразования List<Row> в List<List<object>>
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Метод не удаляет строку, но назначает ей статус на удаление.
        /// Данный статус будет учитываться при удалении строк из google таблицы.
        /// </summary>
        /// <param name="row"></param>
        public void DeleteRow(Row row)
        {
            row.Status = RowStatus.ToDelete;
        }

        /// <summary>
        /// Метод получения ValueRange из строк со статусом ToChange.
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
        /// Метод для получения групп строк для удаления.
        /// </summary>
        /// <returns></returns>
        internal List<List<Row>> GetDeleteRows()
        {
            var deleteGroups = new List<List<Row>>()
            {
                new List<Row>()
            };

            var rowsToDelete = Rows.FindAll(row => row.Status == RowStatus.ToDelete);
            // Данный список нужно зеркалить потому что удаление строк должно происходить с конца таблицы.
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
    }
}
