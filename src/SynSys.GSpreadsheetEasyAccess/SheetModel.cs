﻿using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess
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
        public string Title { get; internal set; } = string.Empty;
        /// <summary>
        /// Id google листа
        /// </summary>
        public string SpreadsheetId { get; internal set; } = string.Empty;
        /// <summary>
        /// Id листа
        /// </summary>
        /// <remarks>
        /// Свойство имеет такое название потому что в uri листа ему
        /// соответствует параметр gid.
        /// </remarks>
        public string Gid { get; internal set; } = string.Empty;
        /// <summary>
        /// Имя листа
        /// </summary>
        public string SpreadsheetTitle { get; internal set; } = string.Empty;
        /// <summary>
        /// Текущий статус.<br/>
        /// <c>Пример: Лист не содержит данных</c>
        /// </summary>
        public string Status { get; internal set; } = string.Empty;
        /// <summary>
        /// Имя ключевого столбца, если он есть.
        /// </summary>
        public string KeyName { get; internal set; } = string.Empty;
        /// <summary>
        /// Режим, по которому определяется работа с листом.
        /// </summary>
        public SheetMode Mode { get; internal set; }
        /// <summary>
        /// Шапка (первая строка листа) если есть
        /// </summary>
        public List<string> Head { get; internal set; }
        /// <summary>
        /// Все строки входящие в данный лист
        /// за исключением шапки
        /// </summary>
        public List<Row> Rows { get; } = new List<Row>();

        /// <summary>
        /// Инициализирует пустой экземпеляр листа готовый для заполнения.
        /// Экземпляр листа нельзя создавать вне библиотеки.
        /// </summary>
        internal SheetModel() { }

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
        /// <param name="rowData">Данные для составленния строки</param>
        public void AddRow(List<string> rowData)
        {
            var number = 1;

            if (Rows.Count > 0)
            {
                number = Rows.Last().Number;
            }

            AddRow(
                number,
                Head.Count,
                rowData
            );
        }

        /// <summary>
        /// Метод не удаляет строку, но назначает ей статус на удаление.
        /// Данный статус будет учитываться при удалении строк из листа в google.
        /// </summary>
        /// <param name="row"></param>
        public void DeleteRow(Row row)
        {
            row.Status = RowStatus.ToDelete;
        }


        /// <summary>
        /// Заполнение листа с созданием строк и ячеек.
        /// </summary>
        /// <param name="data">Данные для формирования листа</param>
        internal void Fill(IList<IList<object>> data)
        {
            var rows = Enumerable.Range(0, data.Count);
            var maxColumns = data.First().Count;

            foreach (var rowIndex in rows)
            {
                if (rowIndex == 0)
                {
                    if (Mode == SheetMode.Simple)
                    {
                        CreateEmptyHead(maxColumns);
                    }
                    else
                    {
                        Head = data[rowIndex].Cast<string>().ToList();
                        continue;
                    }
                }

                var rowData = data[rowIndex].Cast<string>().ToList();
                AddRow(rowIndex, maxColumns, rowData, RowStatus.Original);
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
 
        internal void RenumberRows()
        {
            int number = 1;

            switch (Mode)
            {
                case SheetMode.Simple:
                    number = 1;
                    break;
                case SheetMode.Head:
                    number = 2;
                    break;
                case SheetMode.HeadAndKey:
                    number = 2;
                    break;
            }

            foreach (var row in Rows)
            {
                row.Number = number;
                number++;
            }
        }

        /// <summary>
        /// Сброс статуса всех строк на Original
        /// чтобы после обновления данных в гугл таблице можно было пользоваться
        /// тем же инстансем типа Sheet.
        /// </summary>
        internal void ResetRowStatuses()
        {
            foreach (var row in Rows.FindAll(row => row.Status != RowStatus.Original))
            {
                row.Status = RowStatus.Original;
            }
        }


        /// <summary>
        /// Создание пустой шапки, когда она не нужна для данного листа.
        /// </summary>
        /// <param name="maxLength">Максимальная длина строки</param>
        /// <returns></returns>
        private void CreateEmptyHead(int maxLength)
        {
            Head = new List<string>();

            for (var cellIndex = 0; cellIndex < maxLength; cellIndex++)
            {
                Head.Add(string.Empty);
            }
        }

        /// <summary>
        /// Основной метод для добавления строки в конец листа.
        /// </summary>
        /// <param name="index">Индекс данной строки</param>
        /// <param name="maxLength">Максимальная длина строки</param>
        /// <param name="rowData">Данные для формированния строки</param>
        /// <param name="status">
        /// Статус строки по умолчанию ToAppend.
        /// </param>
        private void AddRow(int index, int maxLength, List<string> rowData, RowStatus status=RowStatus.ToAppend)
        {
            var row = new Row(rowData, maxLength, Head)
            {
                Status = status,
                Number = index + 1
            };

            if (!string.IsNullOrWhiteSpace(KeyName))
            {
                row.Key = row.Cells.Find(cell => cell.Title == KeyName);
            }

            Rows.Add(row);
        }

        /// <summary>
        /// Преобразование List&lt;Row&gt; в List&lt;List&lt;object&gt;&gt;
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
    }
}