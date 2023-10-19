using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    public class NativeSheet : AbstractSheet
    {
        internal NativeSheet() : base()
        {
            _firstRowNumber = 1;
        }

        public void AddColumn()
        {
            var column = new NativeColumn(number: Head.Count + 1)
            {
                Status = ChangeStatus.ToAppend
            };

            Head.Add(column);
            AddCellInAllRows(column);
        }

        public void InsertColumn(int number)
        {
            var column = new NativeColumn(number)
            {
                Status = ChangeStatus.ToInsert
            };
            int columnIndex = column.Number - 1;

            Head.Insert(columnIndex, column);
            InsertCellInAllRows(column, columnIndex);
            RenumberColumnsAfterInserted(column.Number);
            ChangeCellsColumnReferenceAfterInserted(column.Number);
        }


        internal override void CreateHead(List<List<string>> data)
        {
            int maxRowLength = GetMaxRowLength(data);

            for (int columnNumber = 1; columnNumber <= maxRowLength; columnNumber++)
            {
                Head.Add(new NativeColumn(columnNumber));
            }
        }


        protected override void AddRow(int number, int length, IList<string> data, ChangeStatus status=ChangeStatus.ToAppend)
        {
            var row = new Row(data, length, Head)
            {
                Status = status,
                Number = number
            };

            Rows.Add(row);
        }

        protected override int GetMaxRowLength(List<List<string>> data)
        {
            return data.Select(row => row.Count).Max();
        }
    }
}
