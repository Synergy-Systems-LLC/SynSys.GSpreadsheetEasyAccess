using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetGoogleSheetDataAPI
{
    public class Sheet
    {
        public ValueRange Data { get; internal set; }
        public int HorizontalSeparator { get; internal set; }
        public int VerticalSeparator { get; internal set; }
    }
}
