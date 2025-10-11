using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using LiveCharts;
using LiveCharts.Wpf;

namespace DatalogToolMarken
{
    public partial class GraphPage : Page
    {
        public GraphPage(List<DataRecord> data)
        {
            InitializeComponent();

            var temperatures = new ChartValues<double>();
            var labels = new List<string>();

            foreach (var record in data)
            {
                if (double.TryParse(record.Temperature, out double t))
                    temperatures.Add(t);
                labels.Add(record.DateTime);
            }

            // ✅ Explicitly reference LiveCharts.SeriesCollection to avoid PdfSharp conflict
            tempChart.Series = new LiveCharts.SeriesCollection
            {
                new LineSeries
                {
                    Title = "Temperature",
                    Values = temperatures,
                    PointGeometry = DefaultGeometries.Circle,
                    PointGeometrySize = 5
                }
            };

            tempChart.AxisX[0].Labels = labels;
        }
    }
}
