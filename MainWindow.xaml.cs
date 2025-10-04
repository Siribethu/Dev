using ClosedXML.Excel;
using Microsoft.Win32;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Ports;
using System.Windows;
using System.Windows.Threading;

namespace DatalogToolMarken
{
    public partial class MainWindow : Window
    {
        private SerialPort serialPort;
        private ObservableCollection<DataRecord> dataRecords = new ObservableCollection<DataRecord>();

        public MainWindow()
        {
            InitializeComponent();
            dataGridCsv.ItemsSource = dataRecords;
            RefreshPorts();
        }

        // 🔹 Refresh COM Ports (only active ports)
        private void RefreshPorts()
        {
            comboBoxPorts.Items.Clear();
            string[] ports = SerialPort.GetPortNames();

            foreach (var port in ports)
                comboBoxPorts.Items.Add(port);

            if (ports.Length > 0)
                comboBoxPorts.SelectedIndex = 0;
            else
                textBlockStatus.Text = "No active COM ports found.";
        }

        private void buttonRefreshPorts_Click(object sender, RoutedEventArgs e)
        {
            RefreshPorts();
        }

        // 🔹 Serial Connect / Disconnect
        private void buttonSerial_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (serialPort == null || !serialPort.IsOpen)
                {
                    if (comboBoxPorts.SelectedItem == null)
                    {
                        MessageBox.Show("Select a COM port first.");
                        return;
                    }

                    string selectedPort = comboBoxPorts.SelectedItem.ToString();
                    serialPort = new SerialPort(selectedPort, 9600);
                    serialPort.DataReceived += SerialPort_DataReceived;
                    serialPort.Open();

                    textBlockStatus.Text = "Connecting...";
                    buttonSerial.Content = "Connecting...";

                    Dispatcher.InvokeAsync(() =>
                    {
                        textBlockStatus.Text = $"Connected to {selectedPort}";
                        buttonSerial.Content = "Connected";
                    });
                }
                else
                {
                    serialPort.Close();
                    buttonSerial.Content = "Open Serial";
                    textBlockStatus.Text = "Disconnected";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        // 🔹 Serial Data Received from ESP32
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string line = serialPort.ReadLine();
                Dispatcher.Invoke(() =>
                {
                    string[] parts = line.Split(',');

                    if (parts.Length >= 4)
                    {
                        dataRecords.Add(new DataRecord
                        {
                            DateTime = parts[0],
                            Temperature = parts[1],
                            MinTemp = parts[2],
                            Status = parts[3]
                        });
                    }
                });
            }
            catch (Exception) { }
        }

        // 🔹 Download Button
        private void buttonDownload_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Downloading data from ESP32...");
        }

        // 🔹 Save CSV
        private void SaveCSV_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv"
            };

            if (dlg.ShowDialog() == true)
            {
                using (StreamWriter writer = new StreamWriter(dlg.FileName))
                {
                    writer.WriteLine("DateTime,Temperature,MinTemp,Status");
                    foreach (var record in dataRecords)
                        writer.WriteLine($"{record.DateTime},{record.Temperature},{record.MinTemp},{record.Status}");
                }
                MessageBox.Show("CSV file saved successfully!");
            }
        }

        // 🔹 Save Excel
        private void SaveExcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Data");

                ws.Cell(1, 1).Value = "DateTime";
                ws.Cell(1, 2).Value = "Temperature";
                ws.Cell(1, 3).Value = "MinTemp";
                ws.Cell(1, 4).Value = "Status";

                int row = 2;
                foreach (var record in dataRecords)
                {
                    ws.Cell(row, 1).Value = record.DateTime;
                    ws.Cell(row, 2).Value = record.Temperature;
                    ws.Cell(row, 3).Value = record.MinTemp;
                    ws.Cell(row, 4).Value = record.Status;
                    row++;
                }

                wb.SaveAs(dlg.FileName);
                MessageBox.Show("Excel file saved successfully!");
            }
        }

        
        // 🔹 Save PDF (Updated for PDFsharp 6.x)
        private void SavePDF_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    PdfDocument pdf = new PdfDocument();
                    pdf.Info.Title = "Datalog Export";

                    PdfPage page = pdf.AddPage();
                    XGraphics gfx = XGraphics.FromPdfPage(page);

                    // ✅ Use XFontStyleEx instead of XFontStyle
                    XFont titleFont = new XFont("Arial", 14, XFontStyleEx.Bold);
                    XFont textFont = new XFont("Arial", 10, XFontStyleEx.Regular);

                    // ✅ Use XUnit.FromPoint instead of implicit int/double conversions
                    gfx.DrawString("Datalog Export", titleFont, XBrushes.Black,
                        new XRect(XUnit.FromPoint(0), XUnit.FromPoint(20),
                                  page.Width, XUnit.FromPoint(40)),
                        XStringFormats.TopCenter);

                    double y = 70;
                    gfx.DrawString("Date/Time            Temp (°C)     MinTemp     Status",
                        textFont, XBrushes.Black, new XPoint(XUnit.FromPoint(40), XUnit.FromPoint(y)));
                    y += 20;
                    gfx.DrawLine(XPens.Black, XUnit.FromPoint(40), XUnit.FromPoint(y),
                                 page.Width - XUnit.FromPoint(40), XUnit.FromPoint(y));
                    y += 20;

                    foreach (var record in dataRecords)
                    {
                        gfx.DrawString($"{record.DateTime,-20}  {record.Temperature,-10}  {record.MinTemp,-10}  {record.Status}",
                            textFont, XBrushes.Black, new XPoint(XUnit.FromPoint(40), XUnit.FromPoint(y)));
                        y += 20;

                        // Add new page if nearing bottom
                        if (y > page.Height.Point - 40)
                        {
                            page = pdf.AddPage();
                            gfx = XGraphics.FromPdfPage(page);
                            y = 40;
                        }
                    }

                    pdf.Save(dlg.FileName);
                    MessageBox.Show("PDF file saved successfully!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error saving PDF: " + ex.Message);
                }
            }
        }

        // 🔹 Submit Button
        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Data submitted successfully!");
        }
    }

    public class DataRecord
    {
        public string DateTime { get; set; }
        public string Temperature { get; set; }
        public string MinTemp { get; set; }
        public string Status { get; set; }
    }
}
