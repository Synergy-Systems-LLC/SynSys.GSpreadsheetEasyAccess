using System.Collections.Generic;

namespace GSpreadsheetEasyAccess
{
    /// <summary>
    /// Статусы строки. Нужны для определения состояния строки.<br/>
    /// По данному статусу коннектор понимает как работать с конкретной строкой
    /// при отправке данных в google spreadsheet.
    /// </summary>
    public enum RowStatus
    {
        /// <summary>
        /// Оригинальная строка из google таблицы.<br/>
        /// Назначается при загрузке из google spreadsheet.
        /// </summary>
        Original,
        /// <summary>
        /// Строка будет изменена.
        /// </summary>
        ToChange,
        /// <summary>
        /// Строка будет добавлена.
        /// </summary>
        ToAppend,
        /// <summary>
        /// Строка будет удалена.
        /// </summary>
        ToDelete,
    }

    /// <summary>
    /// Тип представляет одину строку данных листа google таблицы.
    /// </summary>
    public class Row
    {
        /// <summary>
        /// Номер, не индекс!
        /// </summary>
        public int Number { get; internal set; }
        /// <summary>
        /// Текущий статус.
        /// </summary>
        public RowStatus Status { get; internal set; } = RowStatus.ToAppend;
        /// <summary>
        /// Все ячейки этой строки.
        /// </summary>
        public List<Cell> Cells { get; internal set; } = new List<Cell>();
        /// <summary>
        /// Ключевая ячейка.
        /// </summary>
        public Cell Key { get; internal set; }

        /// <summary>
        /// Иниализация и заполнение строки таблицы.
        /// </summary>
        /// <remarks>
        /// Строка может быть пустой, при этом будет содержать список пустых ячеек.
        /// </remarks>
        /// <param name="rowData">Данные для заполнения</param>
        /// <param name="maxLength">Назначает максимальную длину строки</param>
        /// <param name="headOfSheet"></param>
        internal Row(List<string> rowData, int maxLength, List<string> headOfSheet)
        {
            for (int cellIndex = 0; cellIndex < maxLength; cellIndex++)
            {
                var value = string.Empty;
                var title = headOfSheet[cellIndex];

                if (cellIndex < rowData.Count)
                {
                    value = rowData[cellIndex];
                }

                Cells.Add(new Cell(value, title, this));
            }
        }

        /// <summary>
        /// Метод для преобразования строки из List&lt;Cell&gt; в List&lt;object&gt;.
        /// Это нужно для подготовки данных к отправке в google таблицу.
        /// </summary>
        /// <returns></returns>
        internal IList<object> GetData()
        {
            var data = new List<object>();

            foreach (var cell in Cells)
            {
                data.Add(cell.Value as object);
            }

            return data;
        }
    }
}