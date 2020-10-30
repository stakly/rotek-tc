using EasyModbus;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace RotekTCCL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string comPort = null;
        private int tempMinVal = 0;
        private int tempMaxVal = 0;
        private int rpmMinVal = 0;
        private int rpmMaxVal = 0;

        private readonly CultureInfo cultureInfo = new CultureInfo("");

        private static readonly TextBox tempMin = new TextBox() { Name = "TempMin", IsEnabled = false };
        private static readonly TextBox tempMax = new TextBox() { Name = "TempMax", IsEnabled = false };
        private static readonly TextBox rpmMin = new TextBox() { Name = "RpmMin", IsEnabled = false };
        private static readonly TextBox rpmMax = new TextBox() { Name = "RpmMax", IsEnabled = false };
        private static readonly Label tempMinLabel = new Label() { Content = "_Температура включения (°C, допустимы десятичные):", Target = tempMin };
        private static readonly Label tempMaxLabel = new Label() { Content = "Т_емпература максимальных оборотов (°C, допустимы десятичные):", Target = tempMax };
        private static readonly Label rpmMinLabel = new Label() { Content = "_Обороты при включении (%):", Target = rpmMin };
        private static readonly Label rpmMaxLabel = new Label() { Content = "_Максимальные обороты (%):", Target = rpmMax };
        private readonly ModbusClient modbusClient = new ModbusClient();

        public MainWindow()
        {
            InitializeComponent();
            GenerateWizardPage1();
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            Next.Content = "Прочитать";
            Prev.IsEnabled = true;
            //StatusText.Text = DateTime.Now.ToString("HH:mm:ss : ") + "Trying to get data from port " + comPort;

            StackPanel.Children.Clear();
            StackPanel.Children.Add(tempMinLabel);
            StackPanel.Children.Add(tempMin);
            StackPanel.Children.Add(tempMaxLabel);
            StackPanel.Children.Add(tempMax);
            StackPanel.Children.Add(rpmMinLabel);
            StackPanel.Children.Add(rpmMin);
            StackPanel.Children.Add(rpmMaxLabel);
            StackPanel.Children.Add(rpmMax);

            try
            {
                //Console.WriteLine(comPort);
                modbusClient.SerialPort = comPort;
                modbusClient.Parity = System.IO.Ports.Parity.None;
                modbusClient.ConnectionTimeout = 500; // msec
                modbusClient.Connect();
                
                ParseDataToTextFields(modbusClient.ReadHoldingRegisters(10, 4));
                
                tempMin.IsEnabled = true;
                tempMax.IsEnabled = true;
                rpmMin.IsEnabled = true;
                rpmMax.IsEnabled = true;

                Write.IsEnabled = true;

                StatusText.Text = String.Format("{0} OK: значения получены", DateTime.Now.ToString("HH:mm:ss"));
            }
            catch (Exception ex)
            {
                StatusText.Text = String.Format("{0} ERR: {1}", DateTime.Now.ToString("HH:mm:ss"), ex.Message);
                tempMin.IsEnabled = false;
                tempMax.IsEnabled = false;
                rpmMin.IsEnabled = false;
                rpmMax.IsEnabled = false;
            }

            return;
        }

        private void Prev_Click(object sender, RoutedEventArgs e)
        {
            GenerateWizardPage1();
            Next.Content = "Далее >";
            //Next.IsEnabled = true;
            Prev.IsEnabled = false;
            Write.IsEnabled = false;
        }

        private void Write_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                tempMinVal = (int)(Single.Parse(tempMin.Text, cultureInfo) * 10);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                tempMaxVal = (int)(Single.Parse(tempMax.Text, cultureInfo) * 10);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            rpmMinVal = int.Parse(rpmMin.Text);
            rpmMaxVal = int.Parse(rpmMax.Text);

            int[] writeValues = new int[8] { tempMinVal, tempMaxVal, rpmMinVal, rpmMaxVal, tempMinVal, tempMaxVal, rpmMinVal, rpmMaxVal };

            try
            {
                //modbusClient.WriteMultipleRegisters(10, writeValues);
                modbusClient.WriteMultipleRegisters(10, writeValues);
                //ParseDataToTextFields(modbusClient.ReadHoldingRegisters(10, 4));
                StatusText.Text = String.Format("{0} OK, записано {1} регистров, получены данные", DateTime.Now.ToString("HH:mm:ss"), writeValues.Length);
            }
            catch (Exception ex)
            {
                StatusText.Text = String.Format("{0} ERR: {1}", DateTime.Now.ToString("HH:mm:ss"), ex.Message);
            }
            return;
        }
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void TemperatureFormatCheck(object sender, TextChangedEventArgs e)
        {
            TextBox tBox = (sender as TextBox);
            Regex regex = new Regex("^\\d+(?:\\.\\d)?$");
            if (regex.IsMatch(tBox.Text))
            {
                tBox.Background = Brushes.White;
                Write.IsEnabled = true;
            }
            else
            {
                tBox.Background = Brushes.Red;
                Write.IsEnabled = false;
            }
        }

        private void RPMFormatCheck(object sender, TextChangedEventArgs e)
        {
            TextBox tBox = (sender as TextBox);
            Regex regex = new Regex("^\\d+$");
            if (regex.IsMatch(tBox.Text))
            {
                tBox.Background = Brushes.White;
                Write.IsEnabled = true;
            }
            else
            {
                tBox.Background = Brushes.Red;
                Write.IsEnabled = false;
            }
        }

        private void ParseDataToTextFields(int[] values)
        {
            int key = 11;
            
            foreach (int val in values)
            {
                switch (key)
                {
                    case 11:
                        tempMinVal = val;
                        break;
                    case 12:
                        tempMaxVal = val;
                        break;
                    case 13:
                        rpmMinVal = val;
                        break;
                    case 14:
                        rpmMaxVal = val;
                        break;
                }
                //Console.WriteLine(key + " : " + val);
                ++key;
            }

            tempMin.TextChanged += TemperatureFormatCheck;
            tempMin.Text = (Single.Parse(tempMinVal.ToString()) / 10).ToString(cultureInfo);

            tempMax.TextChanged += TemperatureFormatCheck;
            tempMax.Text = (Single.Parse(tempMaxVal.ToString()) / 10).ToString(cultureInfo);

            rpmMin.TextChanged += RPMFormatCheck;
            rpmMin.Text = rpmMinVal.ToString();

            rpmMax.TextChanged += RPMFormatCheck;
            rpmMax.Text = rpmMaxVal.ToString();
        }
        private void GenerateWizardPage1()
        {
            StackPanel.Children.Clear();
            StatusText.Text = DateTime.Now.ToString("HH:mm:ss : ") + "Выберите порт и нажмите Далее для получения данных";
            //context = 0;

            var portList = new List<ComPort>();

            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString()).ToList();
                foreach (string s in ports)
                {
                    Match match = Regex.Match(s, "(.*)\\((COM\\d+)\\)");
                    portList.Add(new ComPort(match.Groups[2].Value + " " + match.Groups[1].Value));

                    //Console.WriteLine(match.Groups[2].Value + " " + match.Groups[1].Value);
                }

                portList.Sort();
            }

            foreach (ComPort port in portList)
            {
                if (comPort is null) { comPort = port.ToString().Split(new char[] { ' ' })[0]; }
                RadioButton rb = new RadioButton() { Content = port, IsChecked = port.ToString().StartsWith(comPort + " ") };
                rb.Checked += (sender, args) =>
                {
                    comPort = (sender as RadioButton).Tag.ToString();
                    //Console.WriteLine("Pressed " + comPort);
                };
                rb.Unchecked += (sender, args) => { };
                rb.Tag = port.ToString().Split(new char[] { ' ' })[0];

                StackPanel.Children.Add(rb);
            }
        }
    }

    class ComPort : IComparable<ComPort>
    {
        int _number;
        string _afterNumber;
        string _line;

        public ComPort(string line)
        {
            // Get leading integer.
            Match match = Regex.Match(line, "^COM(\\d+)(.*)$");
            string integer = match.Groups[1].Value;
            this._number = int.Parse(integer);

            // Store string.
            this._afterNumber = match.Groups[2].Value;
            this._line = line;
        }

        public int CompareTo(ComPort other)
        {
            // First compare number.
            int result1 = _number.CompareTo(other._number);
            if (result1 != 0)
            {
                return result1;
            }
            // Second compare part after number.
            return _afterNumber.CompareTo(other._afterNumber);
        }

        public override string ToString()
        {
            return this._line;
        }
    }
}
