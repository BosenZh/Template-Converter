using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Converter.Types;

namespace Converter.Writer
{
    internal class TableInserter
    {
        private readonly WordprocessingDocument doc;

        public TableInserter(WordprocessingDocument doc)
        {
            this.doc = doc;
        }

        public WordprocessingDocument Document
        {
            get { return this.doc; }
        }

        public void Insert(TableData tabletDataObject)
        {
            if (tabletDataObject != null)
            {
                this.InsertTableToBookmarks(tabletDataObject.Bookmark, tabletDataObject);
            }
        }

        protected BorderValues GetBorderStyle(string style)
        {
            switch (style)
            {
                case "Thin":
                    return BorderValues.Single;
                case "Medium": // Mapping Medium to Thick
                case "Thick":
                    return BorderValues.Thick;
                case "Dashed":
                    return BorderValues.Dashed;
                case "Dotted":
                    return BorderValues.Dotted;
                case "Double":
                    return BorderValues.Double;
                default:
                    return BorderValues.None;
            }
        }

        protected uint GetBorderSize(string style)
        {
            switch (style)
            {
                case "Thin":
                    return 6;
                case "Medium":
                    return 12;
                case "Thick":
                    return 18;
                default:
                    return 4;
            }
        }

        protected void SetTableFixedLayout(TableProperties tblProperties)
        {
            TableLayout layout = HelperTools.GetOrCreateChild<TableLayout>(tblProperties);
            layout.Type = TableLayoutValues.Fixed;
        }

        protected SpacingBetweenLines GetZeroSpacing()
        {
            return new SpacingBetweenLines()
            {
                Before = "0",
                After = "0",
                Line = "0",
                LineRule = LineSpacingRuleValues.AtLeast,
            };
        }

        protected void SetCellWidth(double? cellWidthTwips, TableCellProperties tcp)
        {
            TableCellWidth cellWidth = new TableCellWidth()
            {
                Type = TableWidthUnitValues.Dxa,
                Width = ((UInt32Value)cellWidthTwips).ToString(),
            };

            TableCellWidth existingWidth = HelperTools.GetOrCreateChild<TableCellWidth>(tcp);
            existingWidth?.Remove();

            tcp.Append(cellWidth);
        }

        protected Indentation GetIndentation(string justification)
        {
            Indentation indentation = new Indentation();

            switch (justification)
            {
                case "right":
                    indentation.Left = "0";
                    indentation.Right = "100";
                    break;

                case "left":
                    indentation.Left = "100";
                    indentation.Right = "0";
                    break;

                default:
                    indentation.Left = "0";
                    indentation.Right = "0";
                    break;
            }

            indentation.Hanging = "0";
            return indentation;
        }

        private void AppendProperties(Table table)
        {
            EnumValue<BorderValues> val = new EnumValue<BorderValues>(BorderValues.Single);
            const int defaultBorderSize = 0;

            table.AppendChild(
                new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = val, Size = defaultBorderSize },
                        new BottomBorder { Val = val, Size = defaultBorderSize },
                        new LeftBorder { Val = val, Size = defaultBorderSize },
                        new RightBorder { Val = val, Size = defaultBorderSize },
                        new InsideHorizontalBorder { Val = val, Size = defaultBorderSize },
                        new InsideVerticalBorder { Val = val, Size = defaultBorderSize }),
                    new TableJustification { Val = TableRowAlignmentValues.Center }));
        }

        private void SetFontStyles(CellData cellData, RunProperties runProperties)
        {
            if (!string.IsNullOrEmpty(cellData.FontName))
            {
                runProperties.Append(new RunFonts { Ascii = cellData.FontName });
            }

            if (cellData.FontSize > 0)
            {
                runProperties.Append(new FontSize { Val = (cellData.FontSize * 2).ToString() });
            }

            if (!string.IsNullOrEmpty(cellData.FontColor))
            {
                runProperties.Append(new Color { Val = cellData.FontColor });
            }

            if (cellData.Bold)
            {
                runProperties.Append(new Bold());
            }

            if (cellData.Italic)
            {
                runProperties.Append(new Italic());
            }
        }

        private void SetGridSpan(CellData cellData, TableCell cell)
        {
            if (cellData.ColSpan > 1)
            {
                cell.Append(
                    new TableCellProperties(
                        new GridSpan() { Val = cellData.ColSpan }));
            }
        }

        private void SetBackgroundColor(CellData cellData, TableCell cell)
        {
            if (!string.IsNullOrEmpty(cellData.BackgroundColor))
            {
                TableCellProperties cellProperties = new TableCellProperties(new Shading()
                {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = cellData.BackgroundColor.Trim(),
                });
                cell.Append(cellProperties);
            }
        }

        private void SetSpacingAndMargin(ParagraphProperties pPr, TableCellProperties tcp, CellData cellData)
        {
            SpacingBetweenLines existingSpacing = HelperTools.GetOrCreateChild<SpacingBetweenLines>(pPr);
            existingSpacing.Remove();
            pPr.Append(this.GetZeroSpacing());
            pPr.Append(this.GetIndentation(cellData.HorizontalAlignment));
            tcp.Append(this.GetZeroMargin());
        }

        private TableCellMarginDefault GetZeroMargin()
        {
            return new TableCellMarginDefault()
            {
                TopMargin = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                BottomMargin = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                TableCellLeftMargin = new TableCellLeftMargin() { Width = 0, Type = TableWidthValues.Dxa },
                TableCellRightMargin = new TableCellRightMargin() { Width = 0, Type = TableWidthValues.Dxa },
            };
        }

        private TableVerticalAlignmentValues GetTableVerticalAlignment(CellData cellData)
        {
            TableVerticalAlignmentValues verticalAlignmentValue;
            switch (cellData.VerticalAlignment)
            {
                case "center":
                    verticalAlignmentValue = TableVerticalAlignmentValues.Center;
                    break;
                case "top":
                    verticalAlignmentValue = TableVerticalAlignmentValues.Top;
                    break;
                default:
                    verticalAlignmentValue = TableVerticalAlignmentValues.Bottom;
                    break;
            }

            return verticalAlignmentValue;
        }

        private JustificationValues GetJustification(CellData cellData)
        {
            JustificationValues justificationValue;
            switch (cellData.HorizontalAlignment)
            {
                case "center":
                    justificationValue = JustificationValues.Center;
                    break;
                case "right":
                    justificationValue = JustificationValues.Right;
                    break;
                default:
                    justificationValue = JustificationValues.Left;
                    break;
            }

            return justificationValue;
        }

        private void SetRowHeight(CellData cellData, TableRow row)
        {
            const int TwipsAdjustment = 20;
            double? rowHeightTwips = cellData.RowHeight * TwipsAdjustment;
            TableRowHeight rowHeight = new TableRowHeight() { Val = (UInt32Value)rowHeightTwips };
            TableRowProperties rowProperties = this.GetOrCreatePrependedChild<TableRowProperties>(row);
            rowProperties.Append(rowHeight);
        }

        private void ApplyBordersToTableCell(CellData cellData, TableCellProperties tcp)
        {
            TableCellBorders borders = HelperTools.GetOrCreateChild<TableCellBorders>(tcp);
            borders.TopBorder = new TopBorder
            {
                Val = this.GetBorderStyle(cellData.BorderTop.Style),
                Color = cellData.BorderTop.Color,
                Size = this.GetBorderSize(cellData.BorderTop.Style),
                Space = 0,
            };
            borders.BottomBorder = new BottomBorder
            {
                Val = this.GetBorderStyle(cellData.BorderBottom.Style),
                Color = cellData.BorderBottom.Color,
                Size = this.GetBorderSize(cellData.BorderBottom.Style),
                Space = 0,
            };
            borders.LeftBorder = new LeftBorder
            {
                Val = this.GetBorderStyle(cellData.BorderLeft.Style),
                Color = cellData.BorderLeft.Color,
                Size = this.GetBorderSize(cellData.BorderLeft.Style),
                Space = 0,
            };
            borders.RightBorder = new RightBorder
            {
                Val = this.GetBorderStyle(cellData.BorderRight.Style),
                Color = cellData.BorderRight.Color,
                Size = this.GetBorderSize(cellData.BorderRight.Style),
                Space = 0,
            };
        }

        private Run CreateRunWithText(string cellDataValue)
        {
            string textToInsert;
            if (string.IsNullOrEmpty(cellDataValue)
                || cellDataValue.Equals("#N/A", StringComparison.OrdinalIgnoreCase)
                || cellDataValue.Equals("DONE"))
            {
                textToInsert = string.Empty;
            }
            else
            {
                textToInsert = cellDataValue;
            }

            Run run = new Run(new Text(textToInsert));
            return run;
        }

        private bool IsRowEmpty(IEnumerable<CellData> rowGroup)
        {
            return rowGroup.All(cellData =>
                string.IsNullOrEmpty(cellData.Value) ||
                cellData.Value.Equals("#N/A", StringComparison.OrdinalIgnoreCase) ||
                cellData.Value.Equals("DONE", StringComparison.OrdinalIgnoreCase));
        }

        private void TableSetup(
            Table table,
            IOrderedEnumerable<IGrouping<int, CellData>> rows,
            double? totalTableWidthCm)
        {
            double? totalTableWidthTwips = totalTableWidthCm * 1440 / 2.54;
            List<CellData> lastEmptyRowCellData = null;
            TableRow lastNonEmptyRow = null;

            foreach (IGrouping<int, CellData> rowGroup in rows)
            {
                if (this.IsRowEmpty(rowGroup))
                {
                    lastEmptyRowCellData = rowGroup.ToList();
                    continue;
                }

                TableRow row = new TableRow();

                foreach (CellData cellData in rowGroup.OrderBy(c => c.ColumnIndex))
                {
                    Run run = this.CreateRunWithText(cellData.Value);
                    RunProperties runProperties = new RunProperties();
                    this.SetFontStyles(cellData, runProperties);
                    run.PrependChild(runProperties);

                    TableCell cell = new TableCell(new Paragraph(run));

                    this.SetGridSpan(cellData, cell);
                    this.SetBackgroundColor(cellData, cell);

                    Paragraph paragraph = HelperTools.GetOrCreateChild<Paragraph>(cell);

                    ParagraphProperties pPr = this.GetOrCreatePrependedChild<ParagraphProperties>(paragraph);

                    TableCellProperties tcp = this.GetOrCreatePrependedChild<TableCellProperties>(cell);
                    pPr.PrependChild(new KeepLines());
                    pPr.PrependChild(new KeepNext());
                    this.SetSpacingAndMargin(pPr, tcp, cellData);
                    pPr.Append(new Justification() { Val = this.GetJustification(cellData) });
                    tcp.Append(new TableCellVerticalAlignment() { Val = this.GetTableVerticalAlignment(cellData) });
                    double? cellWidthTwips = cellData.ColumnRatio * totalTableWidthTwips;
                    this.SetCellWidth(cellWidthTwips, tcp);
                    this.SetRowHeight(cellData, row);
                    this.ApplyBordersToTableCell(cellData, tcp);  // Apply borders here
                    row.Append(cell);
                }

                table.Append(row);
                lastNonEmptyRow = row;
            }

            // Apply border styles of the last empty row to the last non-empty row
            this.ApplyBorderForEmptyLastRow(lastEmptyRowCellData, lastNonEmptyRow);
        }

        private void ApplyBorderForEmptyLastRow(List<CellData> lastEmptyRowCellData, TableRow lastNonEmptyRow)
        {
            if (lastEmptyRowCellData != null && lastNonEmptyRow != null)
            {
                List<TableCell> lastNonEmptyCells = lastNonEmptyRow.Elements<TableCell>().ToList();
                for (int i = 0; i < lastEmptyRowCellData.Count; i++)
                {
                    if (i < lastNonEmptyCells.Count)
                    {
                        TableCellProperties tcp = lastNonEmptyCells[i].GetFirstChild<TableCellProperties>();
                        if (tcp != null)
                        {
                            this.ApplyBordersToTableCell(lastEmptyRowCellData[i], tcp);
                        }
                    }
                }
            }
        }

        private void InsertToBookmark(string bookmarkName, MainDocumentPart mainPart, Table table)
        {
            BookmarkStart bookmarkStart = mainPart.Document.Body.Descendants<BookmarkStart>()
                                .FirstOrDefault(b => b.Name == bookmarkName);
            BookmarkEnd bookmarkEnd = mainPart.Document.Body.Descendants<BookmarkEnd>()
                .FirstOrDefault(b => b.Id == bookmarkStart?.Id);

            if (bookmarkStart == null)
            {
                Debug.WriteLine($"bookmark{bookmarkName} not found");
                return;
            }

            OpenXmlElement parent = bookmarkStart.Parent;
            if (bookmarkEnd != null)
            {
                bookmarkEnd.Parent.InsertBeforeSelf(table);
                OpenXmlElement nextSibling = bookmarkStart.NextSibling();
                while (nextSibling != null && !nextSibling.Equals(bookmarkEnd))
                {
                    OpenXmlElement temp = nextSibling.NextSibling();
                    nextSibling.Remove();
                    nextSibling = temp;
                }
            }
            else
            {
                parent.InsertAfterSelf(table);
            }
        }

        private void InsertTableToBookmarks(string bookmarkName, TableData tableData)
        {
            MainDocumentPart mainPart = this.doc.MainDocumentPart;

            Table table = new Table();
            this.AppendProperties(table);

            TableProperties tblProperties = HelperTools.GetOrCreateChild<TableProperties>(table);

            this.SetTableFixedLayout(tblProperties);
            IOrderedEnumerable<IGrouping<int, CellData>> rows = tableData.Cells.GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key);
            if (rows == null || !rows.Any())
            {
                Console.WriteLine("Empty Celldata");
                return;
            }

            const double defaultTotalTableWidthCm = 16.5;
            double? totalTableWidthCm = tableData.TableWidth > defaultTotalTableWidthCm
                ? defaultTotalTableWidthCm
                : tableData.TableWidth;

            if (totalTableWidthCm == defaultTotalTableWidthCm)
            {
                TableJustification tableJustification = new TableJustification()
                {
                    Val = TableRowAlignmentValues.Right,
                };
                tblProperties.Append(tableJustification);
            }

            this.TableSetup(table, rows, totalTableWidthCm);

            this.InsertToBookmark(bookmarkName, mainPart, table);
        }

        private T GetOrCreatePrependedChild<T>(OpenXmlCompositeElement parent)
            where T : OpenXmlElement, new()
        {
            T child = parent.GetFirstChild<T>();
            if (child == null)
            {
                child = new T();
                parent.PrependChild(child);
            }

            return child;
        }
    }
}
