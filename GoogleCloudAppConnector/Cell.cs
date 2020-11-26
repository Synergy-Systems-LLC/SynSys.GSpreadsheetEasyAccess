namespace GetGoogleSheetDataAPI
{
    public class Cell
    {
        public Cell(string cellData)
        {
            Value = cellData;
        }

        public string Title { get; set; }
        public string Value { get; set; }
    }
}