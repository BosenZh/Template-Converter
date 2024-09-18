using System.Collections.Generic;

namespace Converter.Types
{
    public class ChartData
    {
        public string Title { get; set; }

        public List<string> Categories { get; set; }

        public List<SeriesData> Series { get; set; }
    }
}
