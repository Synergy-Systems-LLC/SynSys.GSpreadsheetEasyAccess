namespace GetGoogleSheetDataAPI
{
    /// <summary>
    /// Тип представляет ячейку строки.
    /// </summary>
    public class Cell
    {
        private string value;

        /// <summary>
        /// При изменении значения изменяется статус строки в которой находится ячейка на ToChange.
        /// </summary>
        public string Value 
        { 
            get => value;
            set
            {
                this.value = value;
                Host.Status = RowStatus.ToChange;
            }
        }

        /// <summary>
        /// Название столбца в которой находится ячейка.
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// Ссылка на строку в которой эта ячейка располагается.
        /// </summary>
        public Row Host { get; internal set; }

        /// <summary>
        /// Инициализирует ячейку данных таблицы со значением
        /// и ссылкой на строку в которой находится
        /// </summary>
        /// <param name="value"></param>
        /// <param name="title"></param>
        /// <param name="row"></param>
        internal Cell(string value, string title, Row row)
        {
            this.value = value;
            Title = title;
            Host = row;
        }
    }
}