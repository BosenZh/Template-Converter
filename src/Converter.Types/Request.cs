using System;
using System.Collections.Generic;

namespace Converter.Types
{
    public class Request
    {
        public Dictionary<string, string> Text { get; set; } = new Dictionary<string, string>();

        public List<TableData> TableData { get; set; } = new List<TableData>();

        public List<ChartData> ChartData { get; set; } = new List<ChartData>();

        public List<ImageData> ImageData { get; set; } = new List<ImageData>();
    }
}