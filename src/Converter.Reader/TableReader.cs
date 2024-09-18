using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Converter.Types;

namespace Converter.Reader
{
    internal class TableReader
    {
        private readonly SpreadsheetDocument doc;

        public TableReader(SpreadsheetDocument doc)
        {
            this.doc = doc;
        }

        public TableData RetrieveTable(string? worksheetName, string? bookmark, TableRanges tableRanges)
        {
            TableData? tableData = new ();
            Dictionary<(int row, int col), CellData>? cellDictionary = new ();
            int dataStartRow = tableRanges.StartRow;
            int dataStartColumn = tableRanges.StartColumn;
            int dataEndRow = tableRanges.EndRow;
            int dataEndColumn = tableRanges.EndColumn;

            WorkbookPart? workbookPart = this.doc.WorkbookPart;
            WorksheetPart? worksheetPart = GetWorksheetPartByName(workbookPart, worksheetName);
            Worksheet? worksheet = worksheetPart?.Worksheet;

            double? totalTableWidth = 0;
            Dictionary<int, double?>? columnWidths = new ();
            for (int col = dataStartColumn; col <= dataEndColumn; col++)
            {
                double? columnWidth = GetColumnWidth(worksheet, col);
                columnWidths[col] = columnWidth;
                totalTableWidth += columnWidth;
            }

            for (int row = dataStartRow; row <= dataEndRow; row++)
            {
                for (int col = dataStartColumn; col <= dataEndColumn; col++)
                {
                    Cell? cell = GetCell(worksheetPart?.Worksheet, row, col);
                    if (cell == null)
                    {
                        continue;
                    }

                    CellFormat? cellFormat = GetCellFormat(workbookPart, cell);
                    string? numberFormat = GetNumberFormatString(workbookPart, cellFormat);
                    Font? font = GetFont(workbookPart, cellFormat);
                    string? text = GetCellValue(cell, workbookPart);
                    text = ApplyNumberFormat(text, numberFormat);

                    double? columnWidth = columnWidths.ContainsKey(col) ? columnWidths[col] : 0;

                    Alignment? alignment = cellFormat?.Alignment;
                    Border? border = workbookPart?.WorkbookStylesPart?.Stylesheet?.Borders?
                        .Elements<Border>().ElementAt((int)(cellFormat?.BorderId?.Value ?? 0));
                    const double fontSizeModifier = 0.8;
                    const double rowHeightModifier = 0.8;
                    const double lineSpacingModifier = 1.2;
                    double? fontSize = font?.FontSize?.Val?.Value * fontSizeModifier;
                    CellData? cellData = new ()
                    {
                        Value = !string.IsNullOrEmpty(text) ? text : null,
                        RowIndex = row,
                        ColumnIndex = col,
                        BackgroundColor = HexColorNoAlpha(cellFormat?.FillId is not null
                            ? GetColorHex(cellFormat.FillId.Value, workbookPart)
                            : string.Empty),
                        FontName = font?.FontName?.Val,
                        FontSize = fontSize,
                        FontColor = HexColorNoAlpha(font?.Color?.Rgb ?? "FFFAEBC6"),
                        Bold = font?.Bold != null,
                        Italic = font?.Italic != null,
                        ColumnRatio = totalTableWidth > 0 ? columnWidth / totalTableWidth : 0,
                        LineSpacing = fontSize * lineSpacingModifier,
                        RowHeight = GetRowHeight(worksheet, row) * rowHeightModifier,
                        BorderTop = GetTopBorderInfo(border),
                        BorderBottom = GetBottomBorderInfo(border),
                        BorderLeft = GetLeftBorderInfo(border),
                        BorderRight = GetRightBorderInfo(border),
                        HorizontalAlignment = alignment?.Horizontal ?? "center",
                        VerticalAlignment = alignment?.Vertical ?? "center",
                    };
                    BorderColorNoAlpha(cellData);
                    cellDictionary[(row, col)] = cellData;
                }
            }

            // Use the new GetWorksheetMergedCells function
            Dictionary<(int row, int col), (int rowSpan, int colSpan)>? mergedCells
                = GetWorksheetMergedCells(worksheetPart);

            CombineMergedCells(cellDictionary, totalTableWidth, columnWidths, mergedCells);
            RemoveExtraMergedCells(cellDictionary);

            tableData.Cells?.AddRange(cellDictionary.Values);
            tableData.Bookmark = bookmark;
            const double pointToCm = 0.0352778 * 2.54 * 1.75;
            tableData.TableWidth = totalTableWidth * pointToCm;
            return tableData;
        }

        private static Font? GetFont(WorkbookPart? workbookPart, CellFormat? cellFormat) =>
            workbookPart?.WorkbookStylesPart?.Stylesheet?.Fonts?
            .ElementAt((int)(cellFormat?.FontId?.Value ?? 0)) as Font;

        private static void CombineMergedCells(
            Dictionary<(int row, int col), CellData> cellDictionary,
            double? totalTableWidth,
            Dictionary<int, double?> columnWidths,
            Dictionary<(int row, int col), (int rowSpan, int colSpan)> mergedCells)
        {
            foreach ((int row, int col) key in cellDictionary.Keys.ToList())
            {
                if (!mergedCells.ContainsKey(key))
                {
                    continue;
                }

                CellData? cell = cellDictionary[key];
                cell.RowSpan = mergedCells[key].rowSpan;
                cell.ColSpan = mergedCells[key].colSpan;

                // Calculate the column ratio for merged cells
                double? mergedWidth = 0;
                for (int j = 0; j < cell.ColSpan; j++)
                {
                    int colIndex = cell.ColumnIndex + j;
                    if (columnWidths.ContainsKey(colIndex))
                    {
                        mergedWidth += columnWidths[colIndex];
                    }
                }

                cell.ColumnRatio = totalTableWidth > 0 ? mergedWidth / totalTableWidth : 0;

                // Calculate the total height for the merged cell
                double? mergedHeight = 0;
                for (int i = 0; i < cell.RowSpan; i++)
                {
                    int rowIndex = cell.RowIndex + i;
                    double? rowHeight = cellDictionary[(rowIndex, cell.ColumnIndex)].RowHeight;
                    mergedHeight += rowHeight;
                }

                cell.RowHeight = mergedHeight;

                // Combine borders
                CellData? lastColumnCell = cellDictionary[(cell.RowIndex, cell.ColumnIndex + cell.ColSpan - 1)];
                cell.BorderRight = lastColumnCell.BorderRight;

                CellData? lastRowCell = cellDictionary[(cell.RowIndex + cell.RowSpan - 1, cell.ColumnIndex)];
                cell.BorderBottom = lastRowCell.BorderBottom;
            }
        }

        private static void RemoveExtraMergedCells(Dictionary<(int row, int col), CellData> cellDictionary)
        {
            List<(int row, int col)>? cellsToRemove = new ();
            foreach (CellData cell in cellDictionary.Values)
            {
                if (cell.ColSpan > 1 || cell.RowSpan > 1)
                {
                    for (int i = 0; i < cell.RowSpan; i++)
                    {
                        for (int j = 0; j < cell.ColSpan; j++)
                        {
                            if (i != 0 || j != 0)
                            {
                                cellsToRemove.Add((cell.RowIndex + i, cell.ColumnIndex + j));
                            }
                        }
                    }
                }
            }

            foreach ((int, int) key in cellsToRemove)
            {
                cellDictionary.Remove(key);
            }
        }

        private static void BorderColorNoAlpha(CellData cellData)
        {
            cellData.BorderRight.Color = HexColorNoAlpha(cellData.BorderRight?.Color);
            cellData.BorderTop.Color = HexColorNoAlpha(cellData.BorderTop?.Color);
            cellData.BorderLeft.Color = HexColorNoAlpha(cellData.BorderLeft?.Color);
            cellData.BorderBottom.Color = HexColorNoAlpha(cellData?.BorderBottom?.Color);
        }

        private static WorksheetPart? GetWorksheetPartByName(WorkbookPart? workbookPart, string? sheetName)
        {
            Sheet? sheet = workbookPart?.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet?.Id is null)
            {
                return null;
            }

            return workbookPart?.GetPartById(sheet?.Id!) as WorksheetPart;
        }

        private static Cell? GetCell(Worksheet? worksheet, int? rowIndex, int colIndex)
        {
            string? cellReference = GetCellReference(rowIndex, colIndex);
            return worksheet?.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellReference);
        }

        private static double? GetColumnWidth(Worksheet? worksheet, int? colIndex)
        {
            Columns? columns = worksheet?.GetFirstChild<Columns>();
            Column? column = columns?.Elements<Column>().FirstOrDefault(
                c => c.Min is not null
                && c.Max is not null
                && c.Min <= colIndex
                && c.Max >= colIndex)
                ?? new Column();

            const double defaultColumnWidth = 8.43;
            return column?.Width?.Value ??
                worksheet?.SheetFormatProperties?.DefaultColumnWidth ??
                defaultColumnWidth;
        }

        private static CellFormat? GetCellFormat(WorkbookPart? workbookPart, Cell? cell)
        {
            uint? styleIndex = cell?.StyleIndex is not null
                ? cell.StyleIndex.Value
                : 0;
            return workbookPart?.WorkbookStylesPart?.Stylesheet?.CellFormats?.ElementAt((int)styleIndex) as CellFormat;
        }

        private static string? GetNumberFormatString(WorkbookPart? workbookPart, CellFormat? cellFormat)
        {
            uint? numberFormatId = cellFormat?.NumberFormatId?.Value;
            NumberingFormat? numberingFormat = workbookPart?.WorkbookStylesPart?
                .Stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                .FirstOrDefault(nf => nf?.NumberFormatId! == numberFormatId);
            return numberingFormat?.FormatCode?.Value ?? "General";
        }

        private static string? GetColorHex(uint fillId, WorkbookPart? workbookPart)
        {
            const string defaultColor = "FFFFFF";
            Fill? fill = workbookPart?.WorkbookStylesPart?.Stylesheet?.Fills?.ElementAt((int)fillId) as Fill;
            PatternFill? patternFill = fill?.PatternFill;
            if (patternFill != null && patternFill.ForegroundColor != null)
            {
                return patternFill.ForegroundColor?.Rgb?.Value ?? defaultColor;
            }

            return defaultColor;
        }

        private static (int row, int col) GetRowColumnIndex(string? cellReference)
        {
            string? columnReference = new (cellReference?.Where(char.IsLetter).ToArray());
            string? rowReference = new (cellReference?.Where(char.IsDigit).ToArray());

            int columnNumber = 0;
            foreach (char c in columnReference)
            {
                columnNumber = (columnNumber * 26) + c - 'A' + 1;
            }

            int rowNumber = int.Parse(rowReference);
            return (rowNumber, columnNumber);
        }

        private static string? GetCellReference(int? rowIndex, int colIndex)
        {
            string? columnReference = string.Empty;
            while (colIndex > 0)
            {
                int modulo = (colIndex - 1) % 26;
                columnReference = Convert.ToChar('A' + modulo) + columnReference;
                colIndex = (colIndex - modulo) / 26;
            }

            return columnReference + rowIndex;
        }

        private static Dictionary<(int row, int col), (int rowSpan, int colSpan)> GetWorksheetMergedCells(
            WorksheetPart? worksheetPart)
        {
            Dictionary<(int row, int col), (int rowSpan, int colSpan)>? result = new ();

            MergeCells? mergeCells = worksheetPart?.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null)
            {
                return result;
            }

            foreach (MergeCell mergeCell in mergeCells.Elements<MergeCell>())
            {
                string[]? references = mergeCell?.Reference?.Value?.Split(':');
                string? startCell = references?[0];
                string? endCell = references?[1];

                (int startRow, int startCol) = GetRowColumnIndex(startCell);
                (int endRow, int endCol) = GetRowColumnIndex(endCell);

                result[(startRow, startCol)] = (endRow - startRow + 1, endCol - startCol + 1);
            }

            return result;
        }

        private static string? HexColorNoAlpha(string? hexColor)
        {
            const int RGBAColorLength = 8;
            const int AlphaLength = 2;

            return hexColor?.Length == RGBAColorLength
                ? hexColor[AlphaLength..]
                : hexColor;
        }

        private static string? ApplyNumberFormat(string? originalText, string? numberFormat)
        {
            string? formattedText = originalText;

            if (decimal.TryParse(originalText, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal number))
            {
                formattedText = numberFormat switch
                {
                    "General" => originalText,
                    "#,##0" => number.ToString("#,##0", CultureInfo.InvariantCulture),
                    "#,##0.00" => number.ToString("#,##0.00", CultureInfo.InvariantCulture),
                    "0" => number.ToString("0", CultureInfo.InvariantCulture),
                    "0.0%" => (number * 100).ToString("0.0", CultureInfo.InvariantCulture) + "%",
                    "\"$\"#,##0" => FormatCurrency(number, "$", false),
                    "\u0022$\u0022#,##0;\\-\u0022$\u0022#,##0" => FormatCurrency(number, "$", false),
                    "\"$\"#,##0.00" => FormatCurrency(number, "$", true),
                    "\u0022$\u0022#,##0.00;\\-\u0022$\u0022#,##0.00" => FormatCurrency(number, "$", true),
                    "mmmm\\ dd\\,\\ yyyy" => FormatDate(originalText),
                    _ => originalText,
                };
            }
            else if (DateTime.TryParse(originalText, out DateTime date))
            {
                formattedText = numberFormat switch
                {
                    "mmmm\\ dd\\,\\ yyyy" => date.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture),
                    _ => originalText,
                };
            }

            return formattedText;
        }

        private static string? FormatCurrency(decimal number, string? currencySymbol, bool includeDecimals)
        {
            if (number > 0)
            {
                return includeDecimals
                    ? currencySymbol + " " + number.ToString("#,##0.00", CultureInfo.InvariantCulture)
                    : currencySymbol + " " + number.ToString("#,##0", CultureInfo.InvariantCulture);
            }
            else if (number < 0)
            {
                number = 0 - number;
                return includeDecimals
                    ? $"-{currencySymbol} " + number.ToString("#,##0.00", CultureInfo.InvariantCulture)
                    : $"-{currencySymbol} " + number.ToString("#,##0", CultureInfo.InvariantCulture);
            }
            else
            {
                return currencySymbol + "0";
            }
        }

        private static string? FormatDate(string originalText)
        {
            if (double.TryParse(originalText, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
            {
                if (number >= DateTime.MinValue.ToOADate() && number <= DateTime.MaxValue.ToOADate())
                {
                    return DateTime.FromOADate(number).ToString("MMMM d, yyyy", CultureInfo.InvariantCulture);
                }
            }

            return originalText;
        }

        private static double? GetRowHeight(Worksheet? worksheet, int? rowIndex)
        {
            const double defaultRowHeight = 18.75;
            Row? row = worksheet?.Descendants<Row>()
                .FirstOrDefault(r => r.RowIndex is not null && r.RowIndex == rowIndex);
            if (row != null && row.Height is not null)
            {
                return (double)row.Height;
            }
            else if (worksheet?.SheetFormatProperties?.DefaultRowHeight is not null)
            {
                return worksheet.SheetFormatProperties.DefaultRowHeight;
            }

            return defaultRowHeight;
        }

        private static BorderInfo GetBorderInfo(BorderPropertiesType? borderElement)
        {
            if (
                borderElement == null ||
                borderElement.Style == null ||
                borderElement.Style?.Value == BorderStyleValues.None)
            {
                return new BorderInfo { Color = "857874", Style = "Thin" };
            }

            return new BorderInfo
            {
                Color = borderElement.Color?.Rgb?.Value ?? "FFFFFF",
                Style = borderElement.Style?.Value.ToString(),
            };
        }

        private static BorderInfo GetTopBorderInfo(Border? border)
        {
            return border?.TopBorder != null
                ? GetBorderInfo(border.TopBorder)
                : new BorderInfo { Color = "857874", Style = "Thin" };
        }

        private static BorderInfo GetBottomBorderInfo(Border? border)
        {
            return border?.BottomBorder != null
                ? GetBorderInfo(border.BottomBorder)
                : new BorderInfo { Color = "857874", Style = "Thin" };
        }

        private static BorderInfo GetLeftBorderInfo(Border? border)
        {
            return border?.LeftBorder != null
                ? GetBorderInfo(border.LeftBorder)
                : new BorderInfo { Color = "857874", Style = "Thin" };
        }

        private static BorderInfo GetRightBorderInfo(Border? border)
        {
            return border?.RightBorder != null
                ? GetBorderInfo(border.RightBorder)
                : new BorderInfo { Color = "857874", Style = "Thin" };
        }

        private static string? GetCellValue(Cell? cell, WorkbookPart? workbookPart)
        {
            if (cell == null || cell.CellValue == null)
            {
                return null;
            }

            string? value = cell.CellValue.InnerText;

            if (cell.DataType != null)
            {
                value = GetCellValueAccordingToDataType(cell, workbookPart, value);
            }

            return value;
        }

        private static string? GetCellValueAccordingToDataType(Cell? cell, WorkbookPart? workbookPart, string value)
        {
            if (cell?.DataType?.Value == CellValues.SharedString)
            {
                SharedStringTablePart? stringTablePart = workbookPart?.GetPartsOfType<SharedStringTablePart>()
                    .FirstOrDefault();
                if (stringTablePart != null)
                {
                    SharedStringTable? sharedStringTable = stringTablePart.SharedStringTable;
                    if (sharedStringTable != null)
                    {
                        return sharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                }
            }

            if (cell?.DataType?.Value == CellValues.Boolean)
            {
                return value == "0" ? "FALSE" : "TRUE";
            }

            if (cell?.DataType?.Value == CellValues.Date)
            {
                return DateTime.FromOADate(double.Parse(value)).ToString();
            }

            return value;
        }
    }
}