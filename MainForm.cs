using ClosedXML.Excel;
using ScottPlot;
using ScottPlot.Plottables;
using ScottPlot.TickGenerators;
using ScottPlot.TickGenerators.Financial;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;


namespace TemperatureChartWinForms
{
    public class MainForm : Form
    {
        private ScottPlot.WinForms.FormsPlot formsPlot1;
        private ScottPlot.WinForms.FormsPlot formsPlotBar;
        private Button btnUploadCsv;
        private readonly string defaultCsvPath = @"C:\DotNetProject\TemperatureChartApp\DATALOG.csv";

        public MainForm()
        {
            Text = "Temperature Chart - ScottPlot (WinForms)";
            Width = 800;
            Height = 450;

            // Initialize line chart
            formsPlot1 = new ScottPlot.WinForms.FormsPlot { Dock = DockStyle.Fill };
            Controls.Add(formsPlot1);

            // Initialize bar chart
            formsPlotBar = new ScottPlot.WinForms.FormsPlot { Dock = DockStyle.Fill };
            this.Controls.Add(formsPlotBar);


            // ✅ Load chart from default CSV when app starts
            if (File.Exists(defaultCsvPath))
            {
                LoadCsvAndPlot(defaultCsvPath);
                formsPlot1.BringToFront(); // show line chart initially
            }
            else
            {
                MessageBox.Show($"Default CSV file not found:\n{defaultCsvPath}");
            }

            // --- Buttons to switch charts ---

            //var btnBar = new Button { Text = "Bar Chart", Dock = DockStyle.Top, Height = 30 };
            //btnBar.Click += (s, e) => formsPlotBar.BringToFront();
            //var btnLine = new Button { Text = "Line Chart", Dock = DockStyle.Top, Height = 30 };
            //btnLine.Click += (s, e) => formsPlot1.BringToFront();
            //this.Controls.Add(btnLine);




            btnUploadCsv = new Button { Text = "Upload CSV", Dock = DockStyle.Top, Height = 30 };
            btnUploadCsv.Click += BtnUploadCsv_Click;
            Controls.Add(btnUploadCsv);

            // Show line chart initially
            //formsPlot1.BringToFront();

            // Load charts
            //Load += MainForm_Load;
        }
        private void BtnUploadCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "CSV files (*.csv)|*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        LoadCsvAndPlot(ofd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error reading CSV: " + ex.Message);
                    }
                }
            }
        }
        private void LoadCsvAndPlot(string filePath)
        {
            var months = new List<DateTime>();
            var temps = new List<double>();
            var cabTemp = new List<double>();

            // Expecting CSV format: Date,Temperature,setTemp
            foreach (var line in File.ReadLines(filePath).Skip(8)) // skip header rows
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                var parts = line.Split(',');
                if (parts.Length >= 3)
                {
                   if (DateTime.TryParseExact(parts[0].Trim(), "dd/MM/yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                     out DateTime dt) &&                
                     double.TryParse(parts[2].Trim(), out double settemp) &&
                        double.TryParse(parts[1].Trim(), out double cabTemps))
                    {
                       
                        months.Add(dt);
                        temps.Add(settemp);
                        cabTemp.Add(cabTemps);
                    }
                }
            }

            if (months.Count == 0)
            {
                MessageBox.Show("No valid data found in CSV");
                return;
            }

            // === LINE CHART ===
            double[] xs = months.Select(d => d.ToOADate()).ToArray();

       
            double[] ys1 = temps.ToArray();
            double[] ys2 = cabTemp.ToArray();

            formsPlot1.Plot.Clear();
            var scatter1 = formsPlot1.Plot.Add.Scatter(xs, ys1);
            scatter1.Smooth = true;
            scatter1.Label = "Set Temp";
            scatter1.Color = ScottPlot.Colors.Blue;

            // Second scatter (set temp)
            var scatter2 = formsPlot1.Plot.Add.Scatter(xs, ys2);
            scatter2.Label = "Cabinet Temperature";
            scatter2.Color = ScottPlot.Colors.Red;


            formsPlot1.Plot.Axes.DateTimeTicksBottom();
                   
            formsPlot1.Plot.Axes.Bottom.Label.Text = "Date/Time";
            formsPlot1.Plot.Axes.Left.Label.Text = "Temperature (°C)";
            formsPlot1.Plot.Legend.IsVisible = true;
            //// Add reference line
            //var hLine = formsPlot1.Plot.Add.HorizontalLine(25);
            //hLine.LineWidth = 2;
            //hLine.LabelText = "25 °C";
            formsPlot1.Refresh();
        }


    }
}