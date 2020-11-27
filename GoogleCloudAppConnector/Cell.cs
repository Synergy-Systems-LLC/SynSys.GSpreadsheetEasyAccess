namespace GetGoogleSheetDataAPI
{
    public class Cell
    {
        public Cell()
        {
            Value = string.Empty;
        }

        public Cell(string cellData)
        {
            Value = cellData;
        }

        public string Title { get; set; }
        public string Value { get; set; }
    }
}