using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using MG17_Drivers;
using Zaber.Serial.Core;
using System.IO.Ports;

namespace DrugInjection_GUI
{
    public partial class Form1 : Form
    {
        //intialize file name, communication thread and process for com
        private string fileName = "";
        private Thread commun;
        private Process child;
        private string totalTime = "";
        private bool updateProgress = true;
        
        [BrowsableAttribute(false)]
        public static bool CheckForIllegalCrossThreadCalls { get; set; }

        public Form1()
        {
            InitializeComponent();
            SetConsoleCtrlHandler(ParentCheck, true);
        }
        
       //browsw button
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //store the browswed file name and path 
                txtBox_file.Text = openFileDialog1.FileName;
            }
            Console.WriteLine(result);// <-- For debugging use.

        }
        //Run button
        private void btn_run_Click(object sender, EventArgs e)
        {
            //store the browswed file name and path shown in the text box to the variable fileName
            fileName = txtBox_file.Text;

            //if filename is not an empty string, enter the if block
            if (fileName.Length != 0)
            {
                //if file name exists at the path specified, enter the if block
                if (File.Exists(fileName))
                {
                    //create a new thread that runs the recData function, taking file name as input argument
                    /*commun = new Thread(() => recData(fileName));
                    
                    //start the new thread
                    commun.Start();*/

                    if (!backgroundWorker1.IsBusy)
                    {
                        backgroundWorker1.RunWorkerAsync();
                    }
                   
                    
                    // run button should be set to false once it has been pressed
                    btn_run.Enabled = false;
                    // the pause and stop button should now be set to true since the experiment is now run
                    // and the user has the option to stop or pause the experiment
                    btn_pause.Enabled = true;
                    btn_stop.Enabled = true;

                }
                else
                    // if filename does not exist at the specified path, show a pop-up window with error message
                    MessageBox.Show("File Does Not Exist!!!", "Unknown File Exception", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                // if file name text box is empty, show a pop-up window with error message
                MessageBox.Show("Specify Experiment File", "No File Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        /// <summary>
        /// The function runs in a new thread where is executes the main drug injection code
        /// </summary>
        /// <param name="filename"> excel file with the instructions that the user selected using the browse button</param>
        private void recData (string filename)
        {
            // create a new receiver anonymous pipeline for communication between 
            //this GUI and the main drug automation code
            var receiver = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);

            // store a receiver ID for communication purposes
            string receiverID = receiver.GetClientHandleAsString();

            // store the path location of the main drug injection code
            string clientpath = @"C:\Documents and Settings\Admin\Desktop\DrugInjection\AutomatedDrugInjection\AutomatedDrugInjection\bin\Release\AutomatedDrugInjection.exe";
            string clientpath2 = @"C:\Documents and Settings\Admin\Desktop\DrugInjection\TimeCalculation\TimeCalculation\bin\Debug\TimeCalculation.exe";
            // Creating the process info. 
            var startInfo = new ProcessStartInfo(clientpath,receiverID +" "+ "\""+ fileName + "\""+ " " + "true" );

            startInfo.UseShellExecute = false;
    
            // start the process providing all the information needed to run the main code
            child = Process.Start(startInfo);
            //closes the local copy of the anonymousPipeClientStream's object handle
            receiver.DisposeLocalCopyOfClientHandle();

            //while the process is running i.e the main drug injection code is executing.
            while (!child.HasExited)
            {
                // streamReader object receives info from the main code and stores it in a string line
                // message box is used to show the contents of line
                using (StreamReader sr = new StreamReader(receiver))
                {
                    string line;
                    
                    try
                    {
                        if ((line = sr.ReadLine()) != null)
                        {
                            Console.WriteLine("{0}", line);
                            MessageBox.Show(line);
                        }
                    }
                    catch (Exception e) {  }
                }
            }
            // The default controlling of button defined in another thread is not allowed because of thread safety reasons
            Control.CheckForIllegalCrossThreadCalls = false;
            btn_run.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = true;


        }


        //pause button
        private void btn_pause_Click(object sender, EventArgs e)
        {
            btn_pause.Enabled = false;
            btn_resume.Enabled = true;
            updateProgress = false;
            SuspendProcess(child.Id);
        }

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GenerateConsoleCtrlEvent(uint dwCtrlEvent, uint dwProcessGroupId);

        private delegate bool HandlerRoutine(CtrlTypes CtrlType);

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);

        // control messages
        public enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT,
            CTRL_CLOSE_EVENT,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT
        }

        //stop button
        private void btn_stop_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }
            
            //child.Kill();
            //child.Close();
            
            //GenerateConsoleCtrlEvent(Convert.ToUInt32(CtrlTypes.CTRL_C_EVENT) , Convert.ToUInt32(child.Id));
            //child.WaitForExit();
            //commun.Abort();

            btn_run.Enabled = true;
            btn_pause.Enabled = false;
            btn_stop.Enabled = false;
            btn_resume.Enabled = false;
        }

        private bool ParentCheck(CtrlTypes sig)
        {
            return true;
        }

        //resume button
        private void btn_resume_Click(object sender, EventArgs e)
        {
            btn_resume.Enabled = false;
            btn_pause.Enabled = true;
            updateProgress = true;
            ResumeProcess(child.Id);
            
        }

        // the code below was taken from http://stackoverflow.com/questions/71257/suspend-process-in-c-sharp
        // it is used in the pause and resume button to pause or resume the thread that is running the main
        // drug injection code

        [Flags]
        public enum ThreadAccess : int
        {
            TERMINATE = (0x0001),
            SUSPEND_RESUME = (0x0002),
            GET_CONTEXT = (0x0008),
            SET_CONTEXT = (0x0010),
            SET_INFORMATION = (0x0020),
            QUERY_INFORMATION = (0x0040),
            SET_THREAD_TOKEN = (0x0080),
            IMPERSONATE = (0x0100),
            DIRECT_IMPERSONATION = (0x0200)
        }

        [DllImport("kernel32.dll")]
        static extern IntPtr OpenThread(ThreadAccess dwDesiredAccess, bool bInheritHandle, uint dwThreadId);
        [DllImport("kernel32.dll")]
        static extern uint SuspendThread(IntPtr hThread);
        [DllImport("kernel32.dll")]
        static extern int ResumeThread(IntPtr hThread);
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool) ]
        public static extern bool CloseHandle(IntPtr hObject);


        /// <summary>
        /// pauses the process with ID pid
        /// </summary>
        /// <param name="pid"> process ID </param>
        private static void SuspendProcess(int pid)
        {
            var process = Process.GetProcessById(pid);

            if (process.ProcessName == string.Empty)
                return;

            foreach (ProcessThread pT in process.Threads)
            {
                IntPtr pOpenThread = OpenThread(ThreadAccess.SUSPEND_RESUME, false, (uint)pT.Id);

                if (pOpenThread == IntPtr.Zero)
                {
                    continue;
                }
                //FormClosedEventHandler()
                SuspendThread(pOpenThread);
                CloseHandle(pOpenThread);
                
          
            }
        }
        /// <summary>
        /// resume the process with ID pid
        /// </summary>
        /// <param name="pid"> process ID</param>
        public static void ResumeProcess(int pid)
        {
            var process = Process.GetProcessById(pid);

            if (process.ProcessName == string.Empty)
                return;

            foreach (ProcessThread pT in process.Threads)
            {
                IntPtr pOpenThread = OpenThread(ThreadAccess.SUSPEND_RESUME, false, (uint)pT.Id);

                if (pOpenThread == IntPtr.Zero)
                {
                    continue;
                }

                var suspendCount = 0;
                do
                {
                    suspendCount = ResumeThread(pOpenThread);
                } while (suspendCount > 0);

                CloseHandle(pOpenThread);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Stopwatch s = new Stopwatch();
            Stopwatch s2 = new Stopwatch();
            //TimeSpan s1;
            Dictionary<int, string> funcTimes = new Dictionary<int,string>();

            // create a new receiver anonymous pipeline for communication between 
            //this GUI and the main drug automation code
            var receiver_timecalc = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);

            // store a receiver ID for communication purposes
            string receiverID = receiver_timecalc.GetClientHandleAsString();

            // store the path location of the main drug injection code
            string clientpath = @"C:\Documents and Settings\Admin\Desktop\DrugInjection\AutomatedDrugInjection\AutomatedDrugInjection\bin\Release\AutomatedDrugInjection.exe";
            string clientpath2 = @"C:\Documents and Settings\Admin\Desktop\DrugInjection\TimeCalculation\TimeCalculation\bin\Debug\TimeCalculation.exe";
            // Creating the process info. 
            var startInfo = new ProcessStartInfo(clientpath2, receiverID + " " + "\"" + fileName + "\"" + " " + "true");

            startInfo.UseShellExecute = false;

            // start the process providing all the information needed to run the main code
            child = Process.Start(startInfo);
            //closes the local copy of the anonymousPipeClientStream's object handle
            receiver_timecalc.DisposeLocalCopyOfClientHandle();
            StreamReader sr1 = new StreamReader(receiver_timecalc);
            //while the process is running i.e the main drug injection code is executing.
            while (!child.HasExited)
            {
                // streamReader object receives info from the main code and stores it in a string line
                // message box is used to show the contents of line
                //using (StreamReader sr = new StreamReader(receiver_timecalc))
                //{
                    string line;
                    string[] vals;
                    int row;
                    try
                    {
                        if ((line = sr1.ReadLine()) != null)
                        {
                            vals = line.Split(new[] {' '}, 2);
                            row = Int32.Parse(vals[0]);
                            if (!funcTimes.ContainsKey(row) && row != 0)
                                funcTimes.Add(Int32.Parse(vals[0]), vals[1]);
                            else
                                totalTime = Double.Parse(vals[1]).ToString();

                            Console.WriteLine("{0}", line);
                            //MessageBox.Show(totalTime);
                        }
                    }
                    catch (Exception exc) { }
                //}
            }

            // KILL THE PROCESS IF IT DOES NOT KILL ITSELF

            var receiver = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);

            // store a receiver ID for communication purposes
            receiverID = receiver.GetClientHandleAsString();

            // Creating the process info. 
            startInfo = new ProcessStartInfo(clientpath, receiverID + " " + "\"" + fileName + "\"" + " " + "true");

            startInfo.UseShellExecute = false;

            // start the process providing all the information needed to run the main code
            child = Process.Start(startInfo);
            //closes the local copy of the anonymousPipeClientStream's object handle
            receiver.DisposeLocalCopyOfClientHandle();

            //while the process is running i.e the main drug injection code is executing.
            int progress = 0;
            int progress_method = 0;
            double func_time = 0;
            StreamReader sr = new StreamReader(receiver);
            while (!child.HasExited)
            {
                // streamReader object receives info from the main code and stores it in a string line
                // message box is used to show the contents of line

                    string status;
                    int row;
                    string[] vals;
                    int first_char;
                    try
                    {
                        
                        MessageBox.Show("Trying to read at " + s.Elapsed.TotalSeconds.ToString());
                        //first_char = sr.BaseStream.ReadByte();
                        first_char = sr.Read();
                        MessageBox.Show(first_char.ToString());
                        //sr.BaseStream.
                        if (first_char != 0)
                        {
                            if ((status = sr.ReadLine()) != null)
                            {
                                status = (Convert.ToChar(first_char)).ToString() + status;
                                MessageBox.Show("Got status = " + status + " at " + s.Elapsed.TotalSeconds.ToString());
                                if (status.Equals("Start", StringComparison.OrdinalIgnoreCase))
                                {
                                    // start the timer now
                                    MessageBox.Show("Started at " + s.Elapsed.TotalSeconds.ToString());
                                    s.Start();
                                }
                                else
                                {
                                    row = Int32.Parse(status);

                                    if (funcTimes.ContainsKey(row))
                                    {
                                        MessageBox.Show("Got a row in list at " + s.Elapsed.TotalSeconds.ToString());
                                        progress_method = 0;
                                        s2.Restart();
                                        vals = funcTimes[row].Split(new[] { ' ' }, 2);
                                        func_time = Double.Parse(vals[0]);
                                        lblInstruction.Text = vals[1];
                                        MessageBox.Show("Updated the label at " + s.Elapsed.TotalSeconds.ToString());
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exc) { }
                    MessageBox.Show("Got out of the try catch at " + s.Elapsed.TotalSeconds.ToString());

                    
                    if (updateProgress)
                    {  
                        progress = (int)((s.Elapsed.TotalSeconds / Double.Parse(totalTime)) * 100);
                        MessageBox.Show("Updated Experiment Progress at " + s.Elapsed.TotalSeconds.ToString());
                    }
                    backgroundWorker1.ReportProgress(progress, "Total");

                    if (s2.Elapsed.TotalSeconds > func_time / 100)
                    {
                        if (updateProgress)
                            progress_method++;
                        backgroundWorker1.ReportProgress(progress_method, "Function");
                        MessageBox.Show("Updated Function Progress at " + s.Elapsed.TotalSeconds.ToString());
                    }

                    if (backgroundWorker1.CancellationPending)
                    {
                        MessageBox.Show("Cancelling background worker at " + s.Elapsed.TotalSeconds.ToString());
                        e.Cancel = true;
                        backgroundWorker1.ReportProgress(0, "Total");
                        return;
                    }
                
            }

            s.Stop();

            // The default controlling of button defined in another thread is not allowed because of thread safety reasons
            Control.CheckForIllegalCrossThreadCalls = false;
            btn_run.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = true;

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress;
            switch (e.UserState.ToString())
            {
                case "Total":
                    if (e.ProgressPercentage < 100)
                    {
                        progressBar_tlt.Value = e.ProgressPercentage;
                        progress = e.ProgressPercentage;
                    }
                    else
                    {
                        progressBar_tlt.Value = 100;
                        progress = 100;
                    }

                    lblProgress.Text = "Progress: " + progress.ToString() + "%";
                    break;

                case "Function":
                    if (e.ProgressPercentage < 100)
                    {
                        progressBar_mtd.Value = e.ProgressPercentage;
                        progress = e.ProgressPercentage;
                    }
                    else
                    {
                        progressBar_mtd.Value = 100;
                        progress = 100;
                    }

                    lblMethod.Text = "Progress: " + e.ProgressPercentage.ToString() + "%";
                    break;
            }
            
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

    }
}
