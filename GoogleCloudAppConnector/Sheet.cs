using System.Collections.Generic;
using System.Linq;

namespace GetGoogleSheetDataAPI
{
    /// <summary>
    /// Виды таблиц:
    /// 1. Без шапки и без ключа
    /// 2. Шапка (1 строка)
    /// 3. Шапка (1 строка) + Ключ (1 столцеб)
    /// 4. Шапка (>1 строки) + Составной ключ (>1 столбца)
    /// </summary>
    public class Sheet
    {
        public string Title { get; set; }
        public string SpreadsheetId { get; set; }
        public string Gid { get; set; }
        public List<Row> Rows { get; internal set; } = new List<Row>();
        public string Status { get; internal set; }

        internal Sheet() { }

        public void Fill(IList<IList<object>> data)
        {
            var rows = Enumerable.Range(0, data.Count);
            var columns = Enumerable.Range(0, data.First().Count);

            foreach (int rowIndex in rows)
            {
                var rowData = data[rowIndex];
                var rowModel = new Row()
                {
                    Index = rowIndex,
                    Status = RowStatus.Original
                };

                foreach (int cellIndex in columns)
                {
                    var value = string.Empty;

                    if (cellIndex < rowData.Count())
                    {
                        value = data[rowIndex][cellIndex] as string;
                    }

                    rowModel.Cells.Add(new Cell(value));
                }

                Rows.Add(rowModel);
            }
        }
    }
}
