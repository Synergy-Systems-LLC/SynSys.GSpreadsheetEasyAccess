using Newtonsoft.Json;

namespace SynSys.GSpreadsheetEasyAccess.Data
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
        [JsonProperty]
        public string Value 
        { 
            get => value;
            set
            {
                this.value = value;
                ChangeHostStatus();
            }
        }

        /// <summary>
        /// Название столбца в котором находится ячейка.
        /// </summary>
        [JsonProperty]
        public string Title { get; }

        /// <summary>
        /// Ссылка на строку в которой эта ячейка располагается.
        /// </summary>
        [JsonProperty]
        public Row Host { get; set; }


        [JsonConstructor]
        internal Cell() { }

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


        private void ChangeHostStatus()
        {
            // Эта проверка нужна для случая, когда происходит десериализация листа.
            // В этот момент Host не определён, так как располагается после свойства Value.
            // Если переместить свойство Host до Value, то каждый раз при десериализации
            // у всех строк со статусом RowStatus.Original статус будет меняться на RowStatus.ToChange.
            // Это не происходит при обычной работе, потому что Value назначается через конструктор
            // Cell, а десериализация происходит при передаче значения свойству.
            if (Host == null)
            {
                return;
            }

            // Эта проверка нужна для того, чтобы не менять RowStatus.ToAppend.
            // Потому что не важно сколько раз изменятся значение в добавлямой строке,
            // она всё равно будет добавляться в таблицу и статус RowStatus.ToChange будет не корректным.
            if (Host.Status == RowStatus.Original)
            {
                Host.Status = RowStatus.ToChange;
            }
        }
    }
}