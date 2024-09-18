using SpreadsheetLight;

namespace Converter.Reader
{
    internal class GridReader
    {
        private readonly SLDocument doc;

        public GridReader(SLDocument doc)
        {
            this.doc = doc;
        }

        public string ObtainGrid(string nameRange)
        {
            NameRangeHelper nameRangeHelper = new (this.doc);
            int rowIndex = nameRangeHelper.GetTableRanges(nameRange).StartRow;
            int columnIndex = nameRangeHelper.GetTableRanges(nameRange).StartColumn;

            return this.doc.GetCellValueAsString(rowIndex, columnIndex) ?? string.Empty;
        }
    }
}
