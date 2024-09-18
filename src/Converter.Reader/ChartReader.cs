using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using Converter.Types;

namespace Converter.Reader
{
    internal class ChartReader
    {
        private readonly SpreadsheetDocument doc;

        public ChartReader(SpreadsheetDocument doc)
        {
            this.doc = doc;
        }

        public ChartData? RetrieveChart(string chartName)
        {
            IEnumerable<WorksheetPart>? worksheetParts = this.doc.WorkbookPart?.WorksheetParts;
            if (worksheetParts is null)
            {
                return null;
            }

            foreach (WorksheetPart worksheetPart in worksheetParts)
            {
                DrawingsPart? drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart is null)
                {
                    continue;
                }

                foreach (ChartPart chartPart in drawingsPart.ChartParts)
                {
                    Chart? chart = chartPart.ChartSpace.Elements<Chart>().FirstOrDefault();
                    if (chart is null)
                    {
                        continue;
                    }

                    string? titleText = chart.Descendants<Title>().FirstOrDefault()?.ChartText?.RichText?.InnerText;

                    if (titleText is not null && titleText.Equals(chartName, StringComparison.OrdinalIgnoreCase))
                    {
                        return ExtractChartData(chart, titleText);
                    }
                }
            }

            return null;
        }

        private static ChartData? ExtractChartData(Chart chart, string titleText)
        {
            List<string?> categories = new ();
            List<SeriesData>? series = new ();

            ObtainSeriesData(chart.Descendants<LineChartSeries>(), nameof(LineChartSeries), categories, series);
            ObtainSeriesData(chart.Descendants<PieChartSeries>(), nameof(PieChartSeries), categories, series);
            ObtainSeriesData(chart.Descendants<BarChartSeries>(), nameof(BarChartSeries), categories, series);

            if (series.Count == 0)
            {
                Console.WriteLine("No recognized chart series found in the chart part.");
                return null;
            }

            return new ChartData
            {
                Title = titleText,
                Categories = categories,
                Series = series,
            };
        }

        private static void ObtainSeriesData<T>(
            IEnumerable<T> chartSeriesCollection,
            string seriesType,
            List<string?> categories,
            List<SeriesData>? series)
            where T : OpenXmlElement
        {
            foreach (T chartSeries in chartSeriesCollection)
            {
                List<string?> newCategories = ProcessCategoryAxisData(
                    chartSeries.Elements<CategoryAxisData>().FirstOrDefault());

                categories.Clear();
                categories.AddRange(newCategories);
                string? seriesText = chartSeries.Elements<SeriesText>().FirstOrDefault()?.InnerText;
                Values? values = chartSeries.Elements<Values>().FirstOrDefault();
                if (seriesType == nameof(PieChartSeries))
                {
                    series?.Add(ProcessSeries(seriesType, seriesText, values, categories));
                }
                else if (seriesType == nameof(LineChartSeries))
                {
                    series?.Add(ProcessSeries(seriesType, seriesText, values));
                }
                else
                {
                    throw new Exception($"Unrecognized series type: {seriesType}");
                }
            }
        }

        private static SeriesData ProcessSeries(
            string? seriesType,
            string? seriesName,
            Values? valuesElement,
            List<string?>? categories = null)
        {
            List<double> values = new ();
            categories ??= new List<string?>();
            if (valuesElement?.NumberReference?.NumberingCache != null)
            {
                List<NumericPoint>? numericPoints = valuesElement.NumberReference.NumberingCache
                    .Elements<NumericPoint>().ToList();

                // Ensure that the categories list is the same length as the values list
                while (categories.Count < numericPoints.Count)
                {
                    categories.Add(string.Empty);
                }

                for (int i = numericPoints.Count - 1; i >= 0; i--)
                {
                    if (double.TryParse(
                        numericPoints[i].InnerText,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out double parsedValue))
                    {
                        values.Add(parsedValue);
                        continue;
                    }

                    // Remove the category if the value cannot be parsed
                    Console.WriteLine($"Unable to parse value: {numericPoints[i].InnerText}");
                    numericPoints.RemoveAt(i);
                    categories?.RemoveAt(i);
                }
            }

            if (seriesType != nameof(PieChartSeries))
            {
                values.Reverse();
            }

            return new SeriesData { Name = seriesName, Type = seriesType, Values = values };
        }

        private static List<string?> ProcessCategoryAxisData(CategoryAxisData? categoryAxisDataElement)
        {
            return categoryAxisDataElement?.NumberReference?.NumberingCache?
                .Elements<NumericPoint>().Select(np => np?.NumericValue?.Text).ToList()
                ?? categoryAxisDataElement?.StringReference?.StringCache?.Elements<StringPoint>()
                .Select(sp => sp?.InnerText).ToList()
                ?? new List<string?>();
        }
    }
}
