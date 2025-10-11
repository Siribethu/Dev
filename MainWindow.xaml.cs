using ClosedXML.Excel;
using Microsoft.Win32;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace DatalogToolMarken
{
    public partial class MainWindow : Window
    {
        private SerialPort serialPort;
        private ObservableCollection<DataRecord> dataRecords = new ObservableCollection<DataRecord>();
        private DispatcherTimer usbTimer;

        public MainWindow()
        {
            InitializeComponent();
            RefreshPorts();
            StartUsbWatcher();
        }

        // ✅ USB Auto-Detection Watcher
        private void StartUsbWatcher()
        {
            usbTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(3)
            };
            usbTimer.Tick += UsbTimer_Tick;
            usbTimer.Start();
        }

        private void UsbTimer_Tick(object sender, EventArgs e)
        {
            RefreshPorts();
            bool usbFound = false;

            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType == DriveType.Removable && drive.IsReady)
                {
                    textBlockUsbPath.Text = $"USB Detected: {drive.RootDirectory.FullName}";
                    usbFound = true;
                    break;
                }
            }

            if (!usbFound)
                textBlockUsbPath.Text = "No USB detected";
        }

        // ✅ Refresh COM Ports
        private void RefreshPorts()
        {
            comboBoxPorts.Items.Clear();
            string[] ports = SerialPort.GetPortNames();

            foreach (var port in ports)
            {
                try
                {
                    using (SerialPort sp = new SerialPort(port))
                    {
                        sp.Open();
                        comboBoxPorts.Items.Add(port);
                        sp.Close();
                    }
                }
                catch { }
            }

            if (comboBoxPorts.Items.Count > 0)
                comboBoxPorts.SelectedIndex = 0;
            else
                textBlockStatus.Text = "No active COM ports found.";
        }

        private void buttonRefreshPorts_Click(object sender, RoutedEventArgs e)
        {
            RefreshPorts();
        }

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

                    buttonSerial.Content = "Connecting...";
                    textBlockStatus.Text = $"Connecting to {selectedPort}...";

                    Dispatcher.InvokeAsync(() =>
                    {
                        buttonSerial.Content = "Connected";
                        textBlockStatus.Text = $"Connected to {selectedPort}";
                    });
                }
                else
                {
                    serialPort.Close();
                    buttonSerial.Content = "Connect";
                    textBlockStatus.Text = "Disconnected";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

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
            catch { }
        }

        private void buttonDownload_Click(object sender, RoutedEventArgs e)
        {
            if (serialPort == null || !serialPort.IsOpen)
            {
                MessageBox.Show("Please connect to the COM port first!");
                return;
            }

            try
            {
                // Send a request to the Arduino to send data
                serialPort.WriteLine("GET_DATA"); // Make sure Arduino code listens for this command

                // Optionally, inform the user
                textBlockStatus.Text = "Downloading data from ESP32...";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while downloading: " + ex.Message);
            }
        }


        private void buttonGraph_Click(object sender, RoutedEventArgs e)
        {
            var graphPage = new GraphPage(dataRecords.ToList());
            MainFrame.Navigate(graphPage);
        }


        private void buttonSaveCsv_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog { Filter = "CSV Files (*.csv)|*.csv" };
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

        private void buttonSavePdf_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog { Filter = "PDF Files (*.pdf)|*.pdf" };
            if (dlg.ShowDialog() == true)
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont font = new XFont("Arial", 10);
                XFont headerFont = new XFont("Arial", 14, XFontStyleEx.Bold); // fixed for .NET 8

                double y = 40;
                gfx.DrawString("Datalog Export", headerFont, XBrushes.Black,
                    new XRect(0, y, page.Width, 20), XStringFormats.TopCenter);
                y += 40;

                foreach (var record in dataRecords)
                {
                    gfx.DrawString($"{record.DateTime} | {record.Temperature} | {record.MinTemp} | {record.Status}",
                        font, XBrushes.Black, new XPoint(50, y));
                    y += 20;
                }

                pdf.Save(dlg.FileName);
                MessageBox.Show("PDF file saved successfully!");
            }
        }

        private void buttonSaveExcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx" };
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

        private void buttonSubmit_Click(object sender, RoutedEventArgs e)
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
