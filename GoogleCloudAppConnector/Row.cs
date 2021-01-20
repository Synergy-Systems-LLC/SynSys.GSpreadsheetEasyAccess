using System.Collections.Generic;

namespace GetGoogleSheetDataAPI
{
    /// <summary>
    /// Статусы строки. Нужны для определения состояния строки.
    /// По данному статусу коннектор понимает как работать с конкретной строкой
    /// при отправке данных в google spreadsheet.
    /// </summary>
    public enum RowStatus
    {
        /// <summary>
        /// Оригинальная строка.
        /// Назначается при загрузке из google spreadsheet.
        /// </summary>
        Original,
        ToChange,
        ToAppend,
        ToDelete,
    }

    public class Row
    {
        public int Number { get; internal set; }
        public RowStatus Status { get; internal set; } = RowStatus.ToAppend;
        public List<Cell> Cells { get; internal set; } = new List<Cell>();

        /// <summary>
        /// Иниализация и заполнение строки таблицы.
        /// </summary>
        /// <param name="rowData">Данные для заполнения</param>
        /// <param name="maxLength">Назначает максимальную длину строки</param>
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
        /// Метод для преобразования строки из List<Cell> в List<object>.
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