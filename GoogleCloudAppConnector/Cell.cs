namespace GetGoogleSheetDataAPI
{
    public class Cell
    {
        private string _value;

        /// <summary>
        /// При изменении значения изменяет статус строки в которой находится.
        /// </summary>
        public string Value 
        { 
            get => _value;
            set
            {
                _value = value;
                Host.Status = RowStatus.ToChange;
            }
        }

        /// <summary>
        /// Ссылка на строку в которой эта ячейка располагается.
        /// </summary>
        public Row Host { get; internal set; }

        /// <summary>
        /// Инициализирует ячейку данных таблицы со значением
        /// и ссылкой на строку в которой находится
        /// </summary>
        /// <param name="value"></param>
        /// <param name="row"></param>
        internal Cell(string value, Row row)
        {
            _value = value;
            Host = row;
        }
    }
}