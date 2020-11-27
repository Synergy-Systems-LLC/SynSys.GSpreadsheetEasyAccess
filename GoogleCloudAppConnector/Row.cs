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

        public Row(List<string> rowData, int maxLength)
        {
            for (int cellIndex = 0; cellIndex < maxLength; cellIndex++)
            {
                var value = string.Empty;

                if (cellIndex < rowData.Count)
                {
                    value = rowData[cellIndex] as string;
                }

                Cells.Add(new Cell(value));
            }
        }
    }
}