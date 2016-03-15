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
                textBox1.Text = openFileDialog1.FileName;
            }
            Console.WriteLine(result);// <-- For debugging use.

        }
        //Run button
        private void button2_Click(object sender, EventArgs e)
        {

            //store the browswed file name and path shown in the text box to the variable fileName
            fileName = textBox1.Text;

            //if filename is not an empty string, enter the if block
            if (fileName.Length != 0)
            {
                //if file name exists at the path specified, enter the if block
                if (File.Exists(fileName))
                {
                    //create a new thread that runs the recData function, taking file name as input argument
                    commun = new Thread(() => recData(fileName));
                    
                    //start the new thread
                    commun.Start();

                    // run button should be set to false once it has been pressed
                    Button2.Enabled = false;
                    // the pause and stop button should now be set to true since the experiment is now run
                    // and the user has the option to stop or pause the experiment
                    button3.Enabled = true;
                    button4.Enabled = true;

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
            Button2.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = true;


        }


        //pause button
        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button5.Enabled = true;
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
        private void button4_Click(object sender, EventArgs e)
        {
            child.Kill();
            //child.Close();
            
            //GenerateConsoleCtrlEvent(Convert.ToUInt32(CtrlTypes.CTRL_C_EVENT) , Convert.ToUInt32(child.Id));
            //child.WaitForExit();
            commun.Abort();

            Button2.Enabled = true;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
        }

        private bool ParentCheck(CtrlTypes sig)
        {
            return true;
        }

        //resume button
        private void button5_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            button3.Enabled = true;

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


    }
}
