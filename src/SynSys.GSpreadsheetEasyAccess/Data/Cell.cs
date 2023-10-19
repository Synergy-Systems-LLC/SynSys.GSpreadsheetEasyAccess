using Newtonsoft.Json;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Represents a cell of the row.
    /// </summary>
    public class Cell
    {
        private string value;

        /// <summary>
        /// If the row in which this cell is located has the Original status, 
        /// then the row status will change to ToChange while the cell value changes.
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
        /// The name of the column in which the cell is located.
        /// </summary>
        [JsonProperty]
        public AbstractColumn Column { get; internal set; }

        /// <summary>
        /// Link to the row in which this cell is located.
        /// </summary>
        [JsonProperty]
        public Row Host { get; set; }


        [JsonConstructor]
        internal Cell() { }

        /// <summary>
        /// Initializes a table data cell with a value
        /// and a link to the row in which it is located.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="column"></param>
        /// <param name="row"></param>
        internal Cell(string value, AbstractColumn column, Row row)
        {
            this.value = value;
            Column = column;
            Host = row;
        }


        private void ChangeHostStatus()
        {
            // This check is needed when the sheet is being deserialized.
            // At this moment Host is undefined, as it is located after the property Value.
            // If you move the property Host to Value, then every time you deserialize it
            // the status of all rows with the RowStatus.Original status will change to RowStatus.ToChange.
            // This doesn't happen in normal operation because Value is assigned via the Cell constructor
            // and deserialization happens when the value is passed to the property.
            if (Host == null)
            {
                return;
            }

            // This check is needed in order not to change RowStatus.ToAppend.
            // Because no matter how many times the value in the added line changes,
            // it will still be added to the table and the RowStatus.ToChange status will not be correct.
            // TODO добавить описание второго условия
            if (Host.Status == ChangeStatus.Original && Column.Status == ChangeStatus.Original)
            {
                Host.Status = ChangeStatus.ToChange;
            }
        }
    }
}