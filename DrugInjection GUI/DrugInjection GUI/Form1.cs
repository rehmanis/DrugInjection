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
using DrugInjection_GUI;

namespace DrugInjection_GUI
{
    public partial class Form1 : Form
    {
        //intialize file name, communication thread and process for com
        private string fileName = "";
        private Process child;
        private string totalTime = "";
        private bool updateProgress = true;
        Stopwatch s = new Stopwatch();

        [BrowsableAttribute(false)]
        public static bool CheckForIllegalCrossThreadCalls { get; set; }

        public Form1()
        {
            InitializeComponent();
        }

        //browse button
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                //store the browswed file name and path 
                txtBox_file.Text = openFileDialog1.FileName;
            }
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
                    // Background Worker will execute the Time Calculation code and Automation code
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
                    btn_calibrate.Enabled = false;
                    btn_home.Enabled = false;
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

        //pause button
        private void btn_pause_Click(object sender, EventArgs e)
        {
            s.Stop();
            string log_file = Path.GetDirectoryName(fileName) + @"\" + Path.GetFileNameWithoutExtension(fileName) + "_log.txt";
            using (FileStream f = new FileStream(log_file, FileMode.Append, FileAccess.Write))
            using (StreamWriter writer = new StreamWriter(f))
                writer.WriteLine("Experiment Paused at {0}", DateTime.Now);
            btn_pause.Enabled = false;
            btn_resume.Enabled = true;
            updateProgress = false;
            SuspendProcess(child.Id);
        }


        //stop button
        private void btn_stop_Click(object sender, EventArgs e)
        {
            // Terminate the automation code process
            child.Kill();
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }

            string log_file = Path.GetDirectoryName(fileName) + @"\" + Path.GetFileNameWithoutExtension(fileName) + "_log.txt";
            using (FileStream f = new FileStream(log_file, FileMode.Append, FileAccess.Write))
            using (StreamWriter writer = new StreamWriter(f))
                writer.WriteLine("Experiment Stopped at {0}", DateTime.Now);

            btn_run.Enabled = true;
            btn_pause.Enabled = false;
            btn_stop.Enabled = false;
            btn_resume.Enabled = false;
            btn_calibrate.Enabled = true;
            btn_home.Enabled = true;
        }

        //resume button
        private void btn_resume_Click(object sender, EventArgs e)
        {
            s.Start();

            string log_file = Path.GetDirectoryName(fileName) + @"\" + Path.GetFileNameWithoutExtension(fileName) + "_log.txt";
            using (FileStream f = new FileStream(log_file, FileMode.Append, FileAccess.Write))
            using (StreamWriter writer = new StreamWriter(f))
                writer.WriteLine("Experiment Resumed at {0}", DateTime.Now);

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
        [return: MarshalAs(UnmanagedType.Bool)]
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
            //Stopwatch s = new Stopwatch(); // Manages total experiment time

            Dictionary<int, string> funcTimes = new Dictionary<int, string>();

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

            // Object to read incoming data from pipe
            StreamReader sr1 = new StreamReader(receiver_timecalc);

            //while the process is running i.e the main drug injection code is executing.
            while (!child.HasExited)
            {

                string line;

                try
                {
                    if ((line = sr1.ReadLine()) != null)
                    {
                        totalTime = line;
                        Console.WriteLine("{0}", line);
                    }
                }
                catch (Exception exception) { }

            }

            TimeSpan t = TimeSpan.FromSeconds(double.Parse(totalTime));
            string time = string.Format("{0:D2}d:{1:D2}h:{2:D2}m:{3:D3}s",
                t.Days, t.Hours, t.Minutes, t.Seconds);
            Control.CheckForIllegalCrossThreadCalls = false;
            lbl_time.Text = time;
            Control.CheckForIllegalCrossThreadCalls = true;
            // create a pipe for communication with the automation code
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


            int progress = 0; // Stores percentage progress for experiment
            //StreamReader sr = new StreamReader(receiver);

            //while the process is running i.e the main drug injection code is executing.
            while (!child.HasExited)
            {
                // streamReader object receives info from the main code and stores it in a string line
                // message box is used to show the contents of line

                string status;
                using (StreamReader sr = new StreamReader(receiver))
                {
                    try
                    {
                        // Execution hangs at readline until something is received
                        if ((status = sr.ReadLine()) != null)
                        {
                            if (status.Equals("Start", StringComparison.OrdinalIgnoreCase))
                            {
                                // start the timer now
                                s.Start();
                            }

                        }
                    }
                    catch (Exception exception) { }
                }

                if (updateProgress)
                    progress = (int)((s.Elapsed.TotalSeconds / Double.Parse(totalTime)) * 100);
                backgroundWorker1.ReportProgress(progress);

                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    backgroundWorker1.ReportProgress(0);
                    return;
                }

            }

            backgroundWorker1.ReportProgress(100);
            s.Stop();

            // The default controlling of button defined in another thread is not allowed because of thread safety reasons
            Control.CheckForIllegalCrossThreadCalls = false;
            btn_run.Enabled = true;
            btn_calibrate.Enabled = true;
            btn_home.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = true;

        }

        /// <summary>
        /// Updates the progress percentage for both progress bars
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress;

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
         }


        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void btn_calibrate_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Please ensure that the z-slide is above the microplate and the syringe pump is off, before pressing OK ", "Caution",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (result == DialogResult.OK)
            {
                
                ZaberBinaryPort port = new ZaberBinaryPort("COM4");
                port.Open();
                ZaberBinaryDevice slideX, slideY;
                slideX = new ZaberBinaryDevice(port, 1);
                slideY = new ZaberBinaryDevice(port, 2);
                slideY.Home();
                slideY.PollUntilIdle();
                slideX.MoveAbsolute(0);
                slideX.PollUntilIdle();

                port.Close();
            }
        }

        private void btn_home_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Please ensure that the z-slide is above the microplate and the syringe pump is off, before pressing OK ", "Caution",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (result == DialogResult.OK)
            {
                const int MAX_X = 227527;

                ZaberBinaryPort port = new ZaberBinaryPort("COM4");
                port.Open();
                ZaberBinaryDevice slideX, slideY;
                slideX = new ZaberBinaryDevice(port, 1);
                slideY = new ZaberBinaryDevice(port, 2);
                slideY.Home();
                slideY.PollUntilIdle();
                slideX.MoveAbsolute(MAX_X);
                slideX.PollUntilIdle();

                port.Close();
            }
        }

        

    }
}
