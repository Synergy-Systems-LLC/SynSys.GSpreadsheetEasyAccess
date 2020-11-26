using System.Collections.Generic;

namespace GetGoogleSheetDataAPI
{
    public enum RowStatus
    {
        Original,
        ToChange,
        ToAppend,
        ToDelete,
    }

    public class Row
    {
        public int Index { get; set; }
        public RowStatus Status { get; set; } = RowStatus.ToAppend;
        public List<Cell> Cells { get; set; } = new List<Cell>();
    }
}