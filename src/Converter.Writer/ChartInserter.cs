using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using Converter.Types;

namespace Converter.Writer
{
    internal class ChartInserter
    {
        private readonly WordprocessingDocument doc;

        public ChartInserter(WordprocessingDocument doc)
        {
            this.doc = doc;
        }

        public void Insert(ChartData chartDataObject)
        {
            string chartTemplateName = chartDataObject.Title;
            if (chartDataObject != null)
            {
                this.FillChartTemplate(chartTemplateName, chartDataObject);
            }
        }

        private void UpdateStringCache(StringCache cache, IList<string> values)
        {
            cache.RemoveAllChildren();
            cache.Append(new PointCount() { Val = (uint)values.Count });
            for (uint i = 0; i < values.Count; i++)
            {
                cache.Append(new StringPoint() { Index = i, NumericValue = new NumericValue(values[(int)i]) });
            }
        }

        private void UpdateNumberingCache(NumberingCache cache, IList<double> values)
        {
            cache.RemoveAllChildren();
            cache.Append(new FormatCode("General"));
            cache.Append(new PointCount() { Val = (uint)values.Count });
            for (uint i = 0; i < values.Count; i++)
            {
                cache.Append(
                    new NumericPoint() { Index = i, NumericValue = new NumericValue(values[(int)i].ToString()) });
            }
        }

        private void FillChart<T, TSeries>(T chart, ChartData chartData)
            where T : OpenXmlCompositeElement
            where TSeries : OpenXmlCompositeElement, new()
        {
            List<TSeries> existingSeries = chart.Elements<TSeries>().ToList();
            List<SeriesData> filteredSeries = chartData.Series.Where(s => s.Type == typeof(TSeries).Name).ToList();

            for (int i = 0; i < filteredSeries.Count; i++)
            {
                SeriesData seriesData = filteredSeries[i];
                TSeries chartSeries;

                if (i < existingSeries.Count)
                {
                    chartSeries = existingSeries[i];
                }
                else
                {
                    chartSeries = new TSeries();
                    chartSeries.Append(new Index() { Val = (uint)i });
                    chartSeries.Append(new Order() { Val = (uint)i });
                    chartSeries.Append(new SeriesText(new StringReference(new StringCache(
                        new StringPoint() { Index = 0U, NumericValue = new NumericValue(seriesData.Name) }))));
                    chart.Append(chartSeries);
                }

                this.ProcessSeriesData(seriesData, chartSeries);
                this.ProcessCategoryAxisData(chartData, chartSeries);
                this.ProcessValues(seriesData, chartSeries);
            }
        }

        private void ProcessValues<TSeries>(SeriesData seriesData, TSeries chartSeries)
            where TSeries : OpenXmlCompositeElement, new()
        {
            Values values = HelperTools.GetOrCreateChild<Values>(chartSeries);
            NumberReference numberReference = HelperTools.GetOrCreateChild<NumberReference>(values);
            NumberingCache numberingCache = HelperTools.GetOrCreateChild<NumberingCache>(numberReference);
            this.UpdateNumberingCache(numberingCache, seriesData.Values);
        }

        private void ProcessCategoryAxisData<TSeries>(ChartData chartData, TSeries chartSeries)
            where TSeries : OpenXmlCompositeElement, new()
        {
            CategoryAxisData catAxisData = HelperTools.GetOrCreateChild<CategoryAxisData>(chartSeries);
            StringReference stringReferenceCat = HelperTools.GetOrCreateChild<StringReference>(catAxisData);
            StringCache stringCacheCat = HelperTools.GetOrCreateChild<StringCache>(stringReferenceCat);
            this.UpdateStringCache(stringCacheCat, chartData.Categories);
        }

        private void ProcessSeriesData<TSeries>(SeriesData seriesData, TSeries chartSeries)
            where TSeries : OpenXmlCompositeElement, new()
        {
            SeriesText seriesText = HelperTools.GetOrCreateChild<SeriesText>(chartSeries);
            StringReference stringReference = HelperTools.GetOrCreateChild<StringReference>(seriesText);
            StringCache stringCache = HelperTools.GetOrCreateChild<StringCache>(stringReference);
            this.UpdateStringCache(stringCache, new List<string> { seriesData.Name });
        }

        private void FillChartSpaceWithNewData(Chart chart, ChartData chartData)
        {
            PlotArea plotArea = chart.PlotArea;

            if (plotArea.Elements<LineChart>().Any())
            {
                LineChart lineChart = HelperTools.GetOrCreateChild<LineChart>(plotArea);
                this.FillChart<LineChart, LineChartSeries>(lineChart, chartData);
            }

            if (plotArea.Elements<PieChart>().Any())
            {
                PieChart pieChart = HelperTools.GetOrCreateChild<PieChart>(plotArea);
                this.FillChart<PieChart, PieChartSeries>(pieChart, chartData);
            }

            if (plotArea.Elements<BarChart>().Any())
            {
                BarChart barChart = HelperTools.GetOrCreateChild<BarChart>(plotArea);
                this.FillChart<BarChart, BarChartSeries>(barChart, chartData);
            }
        }

        private void FillChartTemplate(string chartTemplateName, ChartData chartData)
        {
            List<ChartPart> chartParts = this.GetChartParts(chartTemplateName);
            if (chartParts == null || chartParts.Count == 0)
            {
                Debug.WriteLine($"Chart template '{chartTemplateName}' not found.");
                return;
            }

            foreach (ChartPart chartPart in chartParts)
            {
                ChartSpace chartSpace = chartPart.ChartSpace;
                Chart chart = chartSpace.Elements<Chart>().FirstOrDefault();

                if (chart == null)
                {
                    Console.WriteLine("No chart found in the chart part.");
                    continue;
                }

                this.FillChartSpaceWithNewData(chart, chartData);
            }
        }

        private List<ChartPart> GetChartParts(string chartTemplateName)
        {
            return this.doc.MainDocumentPart.ChartParts
                .Where(cp => cp.ChartSpace.Elements<Chart>()
                .FirstOrDefault()?.Title?.ChartText?.RichText?.InnerText == chartTemplateName)
                .ToList();
        }
    }
}
