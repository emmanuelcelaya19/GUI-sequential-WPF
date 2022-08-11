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
        public System.Security.SecureString SecurePassword { get; }
        SerialPort serialPort = new SerialPort();
        bool HipotConnection = false;
        string OMNIAPort = "";
        List<test> TestFlow = new List<test>();
        string SerialNumber = "";
        bool AlarmStatus = false;
        bool TestFlowOK = false;
        int TestSatus = 0;
        bool TestStart = false;
        bool TestStopped = false;
        object dummyNode = null;

        public MainWindow()
        {
            InitializeComponent();
        }


        #region Classes
        public class test
        {
            
            public string TestName { get; set; }
            public bool TestEnable { get; set; }
            public double LowLimit { get; set; }
            public double HighLimit { get; set; }
            public string TestStepStatus { get; set; }
        }

        public class TestStepGridShow
        {
            public string ColumnStepName { get; set; }
            public string ColumnStepStatus { get; set; }
        }
        #endregion Classes


        #region Main function
        private void MainFunction()
        {
            if (FnSerialNumberOk())
            {
                Thread ExecuteTest = new Thread(new ThreadStart(FnRunTest)); // Execute Test in a Different Thread
                ExecuteTest.SetApartmentState(ApartmentState.STA);
                Thread EndTest = new Thread(new ThreadStart(FnEndTest)); // Execute Test in a Different Thread

                FnInitialiceTest();
                ExecuteTest.Start();
                EndTest.Start();
            }
        }

        #endregion


        #region Functions


        private bool FnSerialNumberOk()
        {
            SerialNumber = txt_SerialNumber.Text;
            bool CorrectSN = SerialNumber.Contains("JGH");// && SerialNumber.Length>10;

            if (!CorrectSN)
            {
                MessageBox.Show("Validate Unit Serial Number", "Invalid Serial Number!", MessageBoxButton.OK, MessageBoxImage.Error);
                lbl_TestStatus.Content = "Waiting...";
                lbl_TestStatus.Background = Brushes.Gray;
                txt_SerialNumber.Focus();
                txt_SerialNumber.SelectAll();
            }
            return CorrectSN;
        }

        private void FnInitialiceTest()
        {
            TestSatus = 0;
            TestStart = true;
            TestStopped = false;
            lst_TestFlowConfig.IsEnabled = false;
            btn_start.IsEnabled = false;
            txt_SerialNumber.IsEnabled = false;
            DataTestResults.Items.Clear();
            lbl_TestStatus.Content = "RUNNING...";
            lbl_TestStatus.Background = Brushes.Orange;
        }

        private void FnEndTest()
        {
            while (TestStart)
                continue;

            TestStart = false;
            this.Dispatcher.Invoke(() =>
            {
                lst_TestFlowConfig.IsEnabled = true;
                btn_start.IsEnabled = true;
                txt_SerialNumber.IsEnabled = true;
                txt_SerialNumber.Text = "Enter Serial Number";
                txt_SerialNumber.Focus();
                txt_SerialNumber.SelectAll();

                switch (TestSatus)
                {
                    case 0:
                        lbl_TestStatus.Content = "FAIL";
                        lbl_TestStatus.Background = Brushes.Red;
                        break;

                    case 1:
                        lbl_TestStatus.Content = "PASS";
                        lbl_TestStatus.Background = Brushes.Green;
                        break;

                    case 2:
                        lbl_TestStatus.Content = "ABORT";
                        lbl_TestStatus.Background = Brushes.Blue;
                        break;

                    default:
                        lbl_TestStatus.Content = "Waiting...";
                        lbl_TestStatus.Background = Brushes.Coral;
                        break;
                }
            });

        }

        private void FnRunTest()
        {
            bool CriticalStep = false;   // Variable for control "critical steps", test will not continue if a critical step Fails
            int i = 0;
            TestFlow.RemoveAll(test => !test.TestEnable);

            foreach (test test in TestFlow)
            {

                if (TestStopped == true) // if Abort Test
                {
                    TestStopped = false;
                    TestSatus = 2; // if Abort Test
                    DataTestResults.Dispatcher.Invoke(() =>
                    {
                        DataTestResults.Items.RemoveAt(i);
                        DataTestResults.Items.Insert(i, new TestStepGridShow { ColumnStepName = test.TestName, ColumnStepStatus = "ABORTED" });
                    });
                    break;
                }

                DataTestResults.Dispatcher.Invoke(() =>
                {
                    DataTestResults.Items.Insert(i, new TestStepGridShow { ColumnStepName = test.TestName, ColumnStepStatus = "RUNNING" });
                });


                switch (test.TestName)
                {
                    case "CHECK_INTERLOCK":
                        Thread.Sleep(2000);
                        test.TestStepStatus = FnCheckInterlock("RI?\r");
                        CriticalStep = true;
                        break;

                    case "HIPOT_TEST":
                        Thread.Sleep(2000);
                        test.TestStepStatus = FnHipotTest("TEST\r");
                        CriticalStep = true;
                        break;

                    case "HIPOT_RESET":
                        Thread.Sleep(2000);
                        test.TestStepStatus = FnHipotReset("RESET\r");
                        CriticalStep = true;
                        break;

                    default:
                        MessageBox.Show("Invalid <Step Name>, Validate Config File", "Invalid TestStep Added!", MessageBoxButton.OK, MessageBoxImage.Error);
                        TestStopped = true;
                        break;
                }

                if (TestStopped == true) // if Abort Test
                {
                    TestStopped = false;
                    TestSatus = 2; // if Abort Test
                    DataTestResults.Dispatcher.Invoke(() =>
                    {
                        DataTestResults.Items.RemoveAt(i);
                        DataTestResults.Items.Insert(i, new TestStepGridShow { ColumnStepName = test.TestName, ColumnStepStatus = "ABORTED" });
                    });
                    break;
                }

                DataTestResults.Dispatcher.Invoke(() =>
                {
                    DataTestResults.Items.RemoveAt(i);
                    DataTestResults.Items.Insert(i, new TestStepGridShow { ColumnStepName = test.TestName, ColumnStepStatus = test.TestStepStatus });
                });

                if (test.TestStepStatus != "PASS" && CriticalStep == true)
                {
                    TestSatus = 0;
                    break;
                }

                i++;
            }

            if (TestSatus != 2)
            {
                for (int itera = 0; itera < TestFlow.Count; itera++)
                {
                    if (TestFlow[itera].TestStepStatus == "FAIL")
                    {
                        TestSatus = 0;
                        break;
                    }
                    else
                        TestSatus = 1;
                }
            }


            TestStart = false;
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



  



        private void FnInitializeStation()
        {

            HipotConnection = FnConnectSerialPort();
            TestFlowOK = FnReadConfigDocument();
            FnValidatateStationOK();
        }

        private void FnValidatateStationOK()
        {

            if (HipotConnection && TestFlowOK)
            {
                txt_SerialNumber.IsEnabled = true;
                txt_SerialNumber.Focus();
                txt_SerialNumber.SelectAll();
                btn_start.IsEnabled = true;
            }
            else
            {
                // txt_SerialNumber.IsEnabled = false;
                // btn_start.IsEnabled = false;
            }
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
                    Thread.Sleep(500);
                    string responce = serialPort.ReadLine();
                    //serialPort.Close();
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

        private string SendCommand(string CommandMsj)
        {
            string Responce;
            try
            {
                serialPort.WriteLine(CommandMsj);
                Thread.Sleep(500);
                Responce = serialPort.ReadLine();
            }
            catch (Exception es)
            {
                Responce = null;
                MessageBox.Show(es.Message.ToString(), "Communication Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return Responce;
        }

        private bool FnReadConfigDocument()
        {
            TestFlow.Clear();
            bool OkReadingTestFlow = false;

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
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\emman\OneDrive - fceo.mx\Documentos\FCEO\Plenty\scripts\HipotWalalightProject\Config files" + @"\Config.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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
                        //NewTest.LowLimit = (range.Cells[iterator, colLowLimit] as Excel.Range).Value;
                        //NewTest.HighLimit = (range.Cells[iterator, colHighLimit] as Excel.Range).Value;
                        NewTest.TestEnable = Convert.ToBoolean((range.Cells[iterator, colTesEnable] as Excel.Range).Value);
                        TestFlow.Add(NewTest);
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                    iterator++;
                }
                OkReadingTestFlow = true;

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                test NewTest = new test();
                NewTest.TestName = "Error Reading Config File";
                NewTest.TestEnable = false;
                TestFlow.Add(NewTest);
                OkReadingTestFlow = false;
                MessageBox.Show(ex.ToString());
            }


            FnVisualFlow(TestFlow);

            return OkReadingTestFlow;
        }

        private void FnVisualFlow(List<test> TestFlow)
        {
            lst_TestFlowConfig.Items.Clear();
            foreach (test Testing in TestFlow)
            {
                lst_TestFlowConfig.Items.Add(new test { TestName = Testing.TestName, TestEnable = Testing.TestEnable });
            }
        }


        #endregion


        #region Events

        private void lbl_TestStatus_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (TestStart)
            {
                MessageBoxResult result = MessageBox.Show("Abort Test?", "Abort execution Test", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    TestSatus = 2;
                    TestStopped = true;
                }

            }
        }

        private void HipotWalalight_Loaded(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Arrow;
            FnInitializeStation();
            this.Cursor = Cursors.Arrow;
        }

        private void txt_SerialNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                MainFunction();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MainFunction();

        }

        private void lst_TestFlowConfig_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            TestFlowOK = FnReadConfigDocument();
            FnValidatateStationOK();
            this.Cursor = Cursors.Arrow;
        }

        private void lst_TestFlowConfig_MouseEnter(object sender, MouseEventArgs e)
        {
            lst_TestFlowConfig.ToolTip = "<Double Click> for update Test Flow";
        }

        private void btn_Connect_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            HipotConnection = FnConnectSerialPort();
            FnValidatateStationOK();
            this.Cursor = Cursors.Arrow;
        }


        private void HipotWalalight_Closed(object sender, EventArgs e)
        {
            if (serialPort.IsOpen)
            {
                serialPort.Dispose();
                serialPort.Close();
            }
            Console.WriteLine("killing open resources...");
        }

        private void btn_LoadHipot_File_Click(object sender, RoutedEventArgs e)
        {
            Password.Visibility = Visibility.Visible;
            Password.Focus();
            Password.Password = "";


        }

        private void Password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Password.Visibility = Visibility.Hidden;
                string PassWord = Password.Password.ToString();
                if (PassWord == "Plenty")
                {
                    Password.Password = "";
                    Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



                    // Set filter for file extension and default file extension 
                    dlg.DefaultExt = ".png";
                    dlg.Filter = "JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";
                    dlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();

                    // Display OpenFileDialog by calling ShowDialog method 
                    Nullable<bool> result = dlg.ShowDialog();


                    // Get the selected file name and display in a TextBox 
                    if (result == true)
                    {
                        // Open document 


                        string filename = dlg.FileName;
                        if (filename != null && filename.Length > 0)
                        {
                            string messageBoxText = "Upload File: " + filename;
                            string caption = "HIPOT File Upload";
                            MessageBoxButton button = MessageBoxButton.YesNo;
                            MessageBoxImage icon = MessageBoxImage.Warning;
                            MessageBoxResult resultfile;

                            resultfile = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.Yes);
                            if (resultfile == MessageBoxResult.Yes)
                            {
                                //  FnUploadFile(filename);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Wrong File", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Wrong Password", "Password error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
        }

        private void Password_MouseEnter(object sender, MouseEventArgs e)
        {
            Password.ToolTip = "Enter Admin Password for upload Hipot Files";
        }

        #endregion

    }
}
