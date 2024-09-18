namespace Converter.Types
{
    public class CellData
    {
        public string Value { get; set; }

        public int RowIndex { get; set; }

        public int ColumnIndex { get; set; }

        public int RowSpan { get; set; } = 1;

        public int ColSpan { get; set; } = 1;

        public string BackgroundColor { get; set; }

        public string FontName { get; set; }

        public double? FontSize { get; set; }

        public string FontColor { get; set; }

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public string HorizontalAlignment { get; set; }

        public string VerticalAlignment { get; set; }

        public double? ColumnRatio { get; set; }

        public double? LineSpacing { get; set; }

        public double? RowHeight { get; set; }

        public BorderInfo BorderTop { get; set; }

        public BorderInfo BorderBottom { get; set; }

        public BorderInfo BorderLeft { get; set; }

        public BorderInfo BorderRight { get; set; }
    }
}