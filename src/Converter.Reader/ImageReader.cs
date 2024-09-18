using Converter.Types;
using SpreadsheetLight;

namespace Converter.Reader
{
    internal class ImageReader
    {
        private readonly SLDocument doc;

        public ImageReader(SLDocument doc)
        {
            this.doc = doc;
        }

        public ImageData ObtainImage(string nameRange)
        {
            NameRangeHelper nameRangeHelper = new (this.doc);
            int rowIndex = nameRangeHelper.GetTableRanges(nameRange).StartRow;
            int columnIndex = nameRangeHelper.GetTableRanges(nameRange).StartColumn;

            return new ImageData()
            {
                ImageLink = this.doc.GetCellValueAsString(rowIndex, columnIndex),
                BookmarkName = nameRange,
            };
        }
    }
}
