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
        string OMNIAPort = "";
        List<test> TestFlow = new List<test>();
        string SerialNumber = "";
        bool AlarmStatus = false;
        bool TestSatus = false;

        public MainWindow()
        {
            InitializeComponent();
        }


        #region Class
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
        #endregion Class



        #region Functions
        private void main()
        {
            SerialNumber = txt_SerialNumber.Text;
            bool CorrectSN = SerialNumber.Contains("JGH");// && SerialNumber.Length==10;

            if (CorrectSN)
            {
                FnInitialiceTest();
                TestSatus = FnRunTest();
                FnEndTest(TestSatus);
            }
            else
            {
                FnInitialiceTest();
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
            string StepResult = "";
            int i = 0;
            bool stop = false;

            foreach (test RunTest in TestFlow)
            {
                if (RunTest.TestEnable == true)
                {
                    
                    DataTestResults.Items.Insert(i, new TestStep { TestStepName = RunTest.TestName, TestStepStatus = "RUNNING" });

                    switch (RunTest.TestName)
                    {
                        case "CHECK_INTERLOCK":
                            stop = true;    
                            StepResult = FnCheckInterlock("RI?\r");
                            DataTestResults.Items.RemoveAt(i);
                            DataTestResults.Items.Insert(i, new TestStep { TestStepName = RunTest.TestName, TestStepStatus = StepResult });
                            i++;
                            break;
                        case "HIPOT_TEST":
                            StepResult = FnHipotTest("TEST\r");
                            DataTestResults.Items.RemoveAt(i);
                            DataTestResults.Items.Insert(i, new TestStep { TestStepName = RunTest.TestName, TestStepStatus = StepResult });
                            i++;
                            break;
                        case "HIPOT_RESET":
                            StepResult = FnHipotReset("RESET\r");
                            DataTestResults.Items.RemoveAt(i);
                            DataTestResults.Items.Insert(i, new TestStep { TestStepName = RunTest.TestName, TestStepStatus = StepResult });
                            i++;
                            break;
                        default:
                            StepResult = "FAIL";
                            DataTestResults.Items.RemoveAt(i);
                            DataTestResults.Items.Insert(i, new TestStep { TestStepName = RunTest.TestName, TestStepStatus = StepResult });
                            MessageBox.Show("Invalid Test Step from file! Steps must be: <CHECK_INTERLOCK, HIPOT_TEST & HIPOT_RESET/>", "Invalid Test flow", MessageBoxButton.OK, MessageBoxImage.Error);
                            return false;

            
                    }

                }

            if (stop == true && StepResult == "FAIL")
                    break;
            }
;           
            if(DataTestResults.Items.Contains("FAIL"))
            return false;
            else
            return true;
        }


        private string FnCheckInterlock(string command)
        {
            string Stepresult = "FAIL";
            int tries = 2;
            int i = 1;
            string HIPOTResponce;
  

            while (true)
            {
                HIPOTResponce = SendCommand(command);
                if (HIPOTResponce == "1")
                {
                    Stepresult = "PASS";
                    break;
                }
                else if (HIPOTResponce == "NO HIPOT CONNECTION")
                {
                    Stepresult = "FAIL";
                    break;
                }
                else if (i <= tries)
                {
                    string messageBoxText = "Retry Test?";
                    string caption = "NO interlock Detected! Retry?";
                    MessageBoxButton button = MessageBoxButton.YesNo;
                    MessageBoxImage icon = MessageBoxImage.Warning;
                    MessageBoxResult result;

                    result = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.Yes);
                    if (result == MessageBoxResult.No)
                    {
                        Stepresult = "FAIL";
                        break;
                    }
                    i++;
                }
                    break;
            }

            return Stepresult;

        }


        private string FnHipotTest(string command)
        {
            //SendCommand(command);
            return "PASS";
        }


        private string FnHipotReset(string command)
        {
            //SendCommand(command);
            return "PASS";
        }



        private string SendCommand(string SendStrig)
        {
            string resultString = "";
            serialPort.BaudRate = 9600;
            serialPort.Parity = Parity.None;
            serialPort.PortName = "COM2";
//            serialPort.PortName = OMNIAPort;
            serialPort.ReadTimeout = 5;
            try
            {

                serialPort.Open();
                serialPort.WriteLine(SendStrig);
                while (serialPort.BytesToRead < 0) continue;
                string responce = serialPort.ReadLine();
                serialPort.Close();
                resultString = responce;
            }
            catch (Exception)
            {
                resultString = "NO HIPOT CONNECTION";
                MessageBox.Show("NO HIPOT Communication", "No Device Responce", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return resultString;
        }

        private void FnEndTest(bool PassFail)
        {

            if (PassFail)
            {
                lbl_TestStatus.Content = "PASS";
                lbl_TestStatus.Background = Brushes.Green;
            }
            else
            {
                lbl_TestStatus.Content = "FAIL";
                lbl_TestStatus.Background = Brushes.Red;
            }

        }

        private void FnWrongSN()
        {
            MessageBox.Show("Validate Unit Serial Number", "Invalid Serial Number!", MessageBoxButton.OK, MessageBoxImage.Error);
            lbl_TestStatus.Content = "Waiting...";
            lbl_TestStatus.Background = Brushes.Gray;
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
                        OMNIAPort = PortName;
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
                //txt_SerialNumber.IsEnabled = false;
               // btn_start.IsEnabled = false;
                btn_Connect.Visibility = Visibility.Visible;
                lbl_HipotStatus.Content = "Offline";
                lbl_HipotStatus.Background = Brushes.Red;
                lbl_AlarmStatus.Content = "N/A";
                MessageBox.Show("Validate HIPOT connections and click <Connect> button", "HIPOT Communication Fail!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
               // txt_SerialNumber.IsEnabled = true;
                //btn_start.IsEnabled = true;
                btn_Connect.Visibility = Visibility.Hidden;
                lbl_HipotStatus.Content = "Online";
                lbl_HipotStatus.Background = Brushes.Green;
            }
            return Found;
        }


        #endregion


        #region Events

        private void txt_SerialNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                main();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            main();

        }

        private void lbl_TestStatus_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (lbl_TestStatus.Content == "RUNNING")
            {
                string messageBoxText = "Stop Test?";
                string caption = "Test";
                MessageBoxButton button = MessageBoxButton.YesNo;
                MessageBoxImage icon = MessageBoxImage.Warning;
                MessageBoxResult result;

                result = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.Yes);
                if (result == MessageBoxResult.Yes)
                {
                    // StopTest = true;
                }

            }
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


    }
}
