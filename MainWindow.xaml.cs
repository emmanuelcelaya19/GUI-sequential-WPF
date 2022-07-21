using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.IO.Ports;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Threading;

namespace HipotWalalightProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SerialPort serialPort = new SerialPort();
        bool HipotConnection = false;
        List<test> TestFlow = new List<test>();
        string SerialNumber = "";
        bool AlarmStatus = false;
        bool TestSatus = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        #region Events

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            main();

        }

        private void HipotWalalight_Loaded(object sender, RoutedEventArgs e)
        {
            HipotConnection = FnConnectSerialPort();
            FnInitialization();
            TestFlow = FnReadConfigDocument();
            FnVisualFlow(TestFlow);
        }


        private void btn_Connect_Click(object sender, RoutedEventArgs e)
        {
            HipotConnection = FnConnectSerialPort();
        }

        #endregion


        #region Functions

        public class test
        {
            public string TestName { get; set; }
            public bool TestEnable { get; set; }
            public double LowLimit { get; set; }
            public double HighLimit { get; set; }
        }

        public class TestStep
        {
            public string TestStepName { get; set; }
            public string TestStepStatus { get; set; }
        }

        private void main()
        {
            SerialNumber = txt_SerialNumber.Text;
            bool CorrectSN = SerialNumber.Contains("JGH");

            if (CorrectSN)
            {
                FnInitialiceTest();
                TestSatus = FnRunTest();
                FnEndTest(TestSatus);
            }
            else
            {
                FnWrongSN();
            }
        }

        private void FnInitialiceTest()
        {
            DataTestResults.Items.Clear();
            lbl_TestStatus.Content = "RUNNING...";
            lbl_TestStatus.Background = Brushes.Orange;
        }

        private bool FnRunTest()
        {

            foreach (test RunTestStep in TestFlow)
            {
                if (RunTestStep.TestEnable == true)
                {
                    DataTestResults.Items.Add(new TestStep {TestStepName = RunTestStep.TestName,TestStepStatus = "PASS" });

                }
            }
;
            return false;
        }


        private void FnEndTest(bool PassFail)
        {

            if (PassFail)
            {
                lbl_TestStatus.Content = "PASS";
                lbl_TestStatus.Background = Brushes.GreenYellow;
            }
            else
            {
                lbl_TestStatus.Content = "FAIL";
                lbl_TestStatus.Background = Brushes.Red;
            }

        }

        private void FnWrongSN()
        {
            lbl_TestStatus.Content = "Waiting...";
            lbl_TestStatus.Background = Brushes.Gray;
            MessageBox.Show("Validate Unit Serial Number", "Invalid Serial Number!", MessageBoxButton.OK, MessageBoxImage.Error);
            txt_SerialNumber.Focus();
            txt_SerialNumber.SelectAll();
        }

        private List<test> FnReadConfigDocument()
        {
            TestFlow = new List<test>();

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            string TestName = "empty";
            int colTestName = 2;
            int colLowLimit = 3;
            int colHighLimit = 4;
            int colTesEnable = 5;

            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\emman\OneDrive - fceo.mx\Documentos\FCEO\Plenty\scripts\HipotWalalightProject\Config files" + @"\Config.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                

                range = xlWorkSheet.UsedRange;

                int iterator = 5;

                while (TestName != null)
                {
                   
                #pragma warning disable CS8602 // Dereference of a possibly null reference.
                    TestName = (range.Cells[iterator, colTestName] as Excel.Range).Value2;
                    
                    if (TestName != null)
                    {
                        test NewTest = new test();
                        NewTest.TestName = TestName;
                        NewTest.LowLimit = (range.Cells[iterator, colLowLimit] as Excel.Range).Value;
                        NewTest.HighLimit = (range.Cells[iterator, colHighLimit] as Excel.Range).Value;
                        NewTest.TestEnable = Convert.ToBoolean((range.Cells[iterator,colTesEnable] as Excel.Range).Value);
                        TestFlow.Add(NewTest);
                    }
                #pragma warning restore CS8602 // Dereference of a possibly null reference.

                    iterator++;
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                var json = JsonConvert.SerializeObject(TestFlow);

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }



            return TestFlow;
        }

        private void FnVisualFlow(List<test> TestFlow)
        {
            lst_TestFlowConfig.Items.Clear();
            foreach (test Testing in TestFlow)
            {
                lst_TestFlowConfig.Items.Add(new test{TestName = Testing.TestName, TestEnable = Testing.TestEnable});
            }
        }


        private void FnInitialization()
        {
            txt_SerialNumber.Focus();
            txt_SerialNumber.SelectAll();

        }


        private bool FnConnectSerialPort()
        {
            bool Found = false;
            string[] PortNames = SerialPort.GetPortNames();

            serialPort.PortName = "COM";
            serialPort.BaudRate = 9600;
            serialPort.Parity = Parity.None;

            foreach (string PortName in PortNames)
            {
                try
                {
                    serialPort.PortName = PortName;
                    serialPort.Open();
                    serialPort.WriteLine("*IDN?");
                    while (serialPort.BytesToRead < 0) continue;
                    string responce = serialPort.ReadLine();
                    serialPort.Close();
                    if (responce.Contains("OMNIA"))
                    {
                        Found = true;
                        break;
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Port " + PortName + " is No available");
                }

            }
            if (!Found)
            {
                btn_Connect.Visibility = Visibility.Visible;
                lbl_HipotStatus.Content = "Offline";
                lbl_HipotStatus.Background = Brushes.Red;
                lbl_AlarmStatus.Content = "N/A";
                MessageBox.Show("Validate HIPOT connections and click <Connect> button", "HIPOT Communication Fail!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                btn_Connect.Visibility = Visibility.Hidden;
                lbl_HipotStatus.Content = "Online";
                lbl_HipotStatus.Background = Brushes.Green;
            }
            return Found;
        }


        #endregion

        private void lst_TestFlowConfig_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
