using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Converter.Types;
using SpreadsheetLight;

namespace Converter.Reader
{
    public class Program
    {
        public static Request ReadExcel(string filePath, string templatePath)
        {
            List<string> existingBookmarks = GetAllBookmarkNames(filePath);
            Dictionary<string, string> namedRanges = new Dictionary<string, string>();
            Dictionary<string, (string, TableRanges)> tableRanges = new Dictionary<string, (string, TableRanges)>();

            List<ImageData> images = new List<ImageData>();
            List<TableData> tables = new List<TableData>();
            List<ChartData> charts = new List<ChartData>();
            Dictionary<string, string> text = new Dictionary<string, string>();
            List<string> chartTitles = GetCommonChartTitles(filePath, templatePath);

            SLDocument doc = new(filePath);
            NameRangeHelper nameRangeHelper = new(doc);
            List<string> ranges = doc.GetWorksheetNames();
            foreach (SLDefinedName nameRange in doc.GetDefinedNames())
            {
                string name = nameRange.Name;
                string range = nameRange.Text;
                if (!existingBookmarks.Contains(nameRange.Name))
                {
                    continue;
                }
                
                if (range.Contains(":"))
                {
                    string worksheet = nameRangeHelper.GetWorksheetOfNamedRange(name);
                    TableRanges tableRange = nameRangeHelper.GetTableRanges(name);
                    tableRanges[nameRange.Name] = (worksheet,tableRange);
                }
                else
                {
                    namedRanges[nameRange.Name] = range;
                }
            }

            GridReader gridReader = new GridReader(doc);
            ImageReader imageReader = new ImageReader(doc);
            
            foreach (string range in namedRanges.Keys)
            {
                string context = gridReader.ObtainGrid(range);
                if (IsImageLink(context))
                {
                    images.Add(imageReader.ObtainImage(range));
                }
                else
                {
                    text[range] = context;
                }
            }

            SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false);
            TableReader tableReader = new TableReader(document);
            ChartReader chartReader = new ChartReader(document);
            foreach (string range in tableRanges.Keys)
            {
                string worksheet = tableRanges[range].Item1;
                TableRanges tableRange = tableRanges[range].Item2;
                tables.Add(tableReader.RetrieveTable(worksheet, range, tableRange));
            }
            
            foreach (string chart in chartTitles)
            {
                ChartData cd = chartReader.RetrieveChart(chart);
                if (cd != null)
                {
                    charts.Add(cd);
                }
            }

            Request request = new ()
            {
                Text = text,
                ChartData = charts,
                TableData = tables,
                ImageData = images,
            };

            return request;
        }

        public static List<string> GetCommonChartTitles(string excelFilePath, string wordFilePath)
        {
            // Get chart titles from Excel and Word documents
            List<string> excelChartTitles = GetExcelChartTitles(excelFilePath);
            List<string> wordChartTitles = GetWordChartTitles(wordFilePath);

            // Find common chart titles (case-insensitive comparison)
            List<string> commonTitles = excelChartTitles
                .Intersect(wordChartTitles, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return commonTitles;
        }

        public static List<string> GetWordChartTitles(string filePath)
        {
            List<string> chartTitles = new List<string>();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                // Iterate through all ChartParts in the Word document
                foreach (ChartPart chartPart in wordDoc.MainDocumentPart.ChartParts)
                {
                    // Get the chart title if available
                    if (chartPart.ChartSpace.Descendants<Title>().Count() > 0)
                    {
                        foreach (Title title in chartPart.ChartSpace.Descendants<Title>())
                        {
                            if (title.ChartText != null)
                            {
                                var chartTitleText = title.ChartText.InnerText;
                                chartTitles.Add(chartTitleText);
                            }
                        }
                    }
                }
            }

            return chartTitles;
        }

        public static List<string> GetExcelChartTitles(string filePath)
        {
            List<string> chartTitles = new List<string>();

            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(filePath, false))
            {
                // Iterate through all the worksheet parts
                foreach (WorksheetPart worksheetPart in spreadsheetDoc.WorkbookPart.WorksheetParts)
                {
                    // Iterate through all DrawingsPart to get chart parts
                    if (worksheetPart.DrawingsPart != null)
                    {
                        foreach (ChartPart chartPart in worksheetPart.DrawingsPart.ChartParts)
                        {
                            // Get the chart title if available
                            if (chartPart.ChartSpace.Descendants<Title>().Count() > 0)
                            {
                                foreach (Title title in chartPart.ChartSpace.Descendants<Title>())
                                {
                                    if (title.ChartText != null)
                                    {
                                        var chartTitleText = title.ChartText.InnerText;
                                        chartTitles.Add(chartTitleText);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return chartTitles;
        }

        private static bool IsImageLink(string url)
        {
            // Define a regex pattern for typical image file extensions
            string pattern = @"^.*\.(jpg|jpeg|png|gif|bmp|tiff|tif|webp|svg)$";

            // Check if the URL ends with an image extension (case-insensitive)
            return Regex.IsMatch(url, pattern, RegexOptions.IgnoreCase);
        }

        private static List<string> GetAllBookmarkNames(string filePath)
        {
            List<string> bookmarkNames = new List<string>();

            // Open the Word document using OpenXML
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                // Get bookmarks from the main document part
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                if (mainPart != null || mainPart.Document.Body != null)
                {
                    GetBookmarksFromPart(mainPart.Document.Body!, bookmarkNames);
                }

                // Get bookmarks from headers
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    GetBookmarksFromPart(headerPart.Header, bookmarkNames);
                }

                // Get bookmarks from footers
                foreach (var footerPart in mainPart.FooterParts)
                {
                    GetBookmarksFromPart(footerPart.Footer, bookmarkNames);
                }

                // Get bookmarks from footnotes
                if (mainPart.FootnotesPart != null)
                {
                    GetBookmarksFromPart(mainPart.FootnotesPart.Footnotes, bookmarkNames);
                }

                // Get bookmarks from endnotes
                if (mainPart.EndnotesPart != null)
                {
                    GetBookmarksFromPart(mainPart.EndnotesPart.Endnotes, bookmarkNames);
                }

                // Get bookmarks from comments
                if (mainPart.WordprocessingCommentsPart != null)
                {
                    GetBookmarksFromPart(mainPart.WordprocessingCommentsPart.Comments, bookmarkNames);
                }
            }

            return bookmarkNames;
        }

        // Helper method to retrieve bookmarks from a specific document part
        private static void GetBookmarksFromPart(OpenXmlElement partElement, List<string> bookmarkNames)
        {
            if (partElement != null)
            {
                foreach (BookmarkStart bookmarkStart in partElement.Descendants<BookmarkStart>())
                {
                    if (bookmarkStart is not null && bookmarkStart.Name is not null)
                    {
                        bookmarkNames.Add(bookmarkStart.Name!);
                    }
                }
            }
        }


        private static string GetExpenditurePerUnitComment(
            double firstIncreaseRate,
            double contributionPerUnit,
            double expenditurePerUnit,
            double longTermContribution)
        {
            if (firstIncreaseRate > longTermContribution && expenditurePerUnit > contributionPerUnit)
            {
                return "The forecasted average expenditures per unit exceed the contribution " +
                    "per unit in the next fiscal year. The contribution increases recommended for future years" +
                    " will address the resultant shortfall.";
            }
            else if (firstIncreaseRate > longTermContribution && expenditurePerUnit <= contributionPerUnit)
            {
                return "Although the average expenditure per unit is forecast to be less than" +
                    " the average contribution per unit, the timing of cash flows in the proposed expenditure plan" +
                    " still necessitates an increase in annual reserve fund contributions. The contribution increase" +
                    " required is reflected in our proposed funding plan.";
            }
            else if (firstIncreaseRate <= longTermContribution && expenditurePerUnit > contributionPerUnit)
            {
                return "Although the forecasted average expenditures per unit exceed the" +
                    " contribution per unit in the next fiscal year, the Corporation's current financial resources" +
                    ", budgeted reserve fund allocations, and anticipated interest earnings are sufficient to avoid" +
                    " increasing contributions above the rate of inflation.";
            }

            return "The average expenditures per unit are less than the contribution" +
                    " per unit in the next fiscal year.";
        }

        private static string FromExcelSerialDate(int serialDate)
        {
            DateTime date = new DateTime(1900, 1, 1).AddDays(serialDate - 2);
            return date.ToString("MMMM d, yyyy");
        }

        private static string FromExcelSerialDateUsingSlash(int serialDate)
        {
            DateTime date = new DateTime(1900, 1, 1).AddDays(serialDate - 2);
            return date.ToString("MM/dd/yyyy");
        }
    }
}