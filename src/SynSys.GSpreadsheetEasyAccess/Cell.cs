namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Тип представляет ячейку строки.
    /// </summary>
    public class Cell
    {
        private string value;

        /// <summary>
        /// Если строка, в которой находится данная ячейка имеет статус Original, то
        /// при изменении значения ячейки, изменится статус строки на ToChange.
        /// </summary>
        public string Value 
        { 
            get => value;
            set
            {
                this.value = value;
                // Данная проверка нужна для того, чтобы не менять RowStatus.ToAppend.
                // Потому что не важно сколько раз изменятся значения в добавлямой строке,
                // она всё равно будет добавляться в таблицу.
                // Изменения могут происходить только в существующей строке таблицы.
                if (Host.Status == RowStatus.Original)
                {
                    Host.Status = RowStatus.ToChange;
                }
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