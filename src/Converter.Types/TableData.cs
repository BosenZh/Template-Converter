using System.Collections.Generic;

namespace Converter.Types
{
    public class TableData
    {
        public string Bookmark { get; set; }

        public double? TableWidth { get; set; }

        public List<CellData> Cells { get; set; } = new List<CellData>();
    }
}