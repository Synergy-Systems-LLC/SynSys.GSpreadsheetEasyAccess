using Newtonsoft.Json;
using System.Collections.Generic;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Determines the state of the row.<br/>
    /// The GCPApplication understands how to work with a specific row (based on this status)
    /// when sending data to Google spreadsheet.
    /// </summary>
    public enum RowStatus
    {
        /// <summary>
        /// Original row from Google spreadsheet.<br/>
        /// Assigned when loading from Google spreadsheet.
        /// </summary>
        Original,
        /// <summary>
        /// The row will be changed.
        /// </summary>
        ToChange,
        /// <summary>
        /// The row will be added.
        /// </summary>
        ToAppend,
        /// <summary>
        /// The row will be deleted.
        /// </summary>
        ToDelete,
    }

    /// <summary>
    /// The type represents one row of one Google spreadsheet sheet.
    /// </summary>
    public class Row
    {
        /// <summary>
        /// Number, not index!
        /// </summary>
        [JsonProperty]
        public int Number { get; internal set; }

        /// <summary>
        /// Current status.
        /// </summary>
        [JsonProperty]
        public RowStatus Status { get; internal set; } = RowStatus.ToAppend;

        /// <summary>
        /// All cells in this row.
        /// </summary>
        [JsonProperty]
        public List<Cell> Cells { get; internal set; } = new List<Cell>();

        /// <summary>
        /// Key cell.
        /// </summary>
        [JsonProperty]
        public Cell Key { get; internal set; }


        [JsonConstructor]
        internal Row() { }

        /// <summary>
        /// Initialization and filling of a spreadsheet row.
        /// </summary>
        /// <remarks>
        /// The row can be empty, and it still will contain a list of empty cells.
        /// </remarks>
        /// <param name="rowData">Data to fill</param>
        /// <param name="maxLength">Assigns the maximum length of a row</param>
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
        /// Row conversion from List&lt;Cell&gt; to List&lt;object&gt;.
        /// This is necessary to prepare data for sending to Google spreadsheet.
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