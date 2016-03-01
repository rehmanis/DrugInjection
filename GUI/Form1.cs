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

namespace DrugInjection_GUI
{
    public partial class Form1 : Form
    {
        private string fileName = "";
        private Thread commun;
        private Process child;

        [BrowsableAttribute(false)]
        public static bool CheckForIllegalCrossThreadCalls { get; set; }

        public Form1()
        {
            InitializeComponent();
        }
        
       
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox1.Text = openFileDialog1.FileName;
            }
            Console.WriteLine(result);// <-- For debugging use.

        }
        //Run button
        private void button2_Click(object sender, EventArgs e)
        {
            

            fileName = textBox1.Text;
            if (fileName.Length != 0)
            {

                if (File.Exists(fileName))
                {
                    commun = new Thread(() => recData(fileName));
                    //MessageBox.Show("After thread");
                    commun.Start();
                    Button2.Enabled = false;
                    button3.Enabled = true;
                    button4.Enabled = true;

                    //timer1.Start();
                    //timer1.Enabled = true;

                }
                else
                    MessageBox.Show("File Does Not Exist!!!", "Unknown File Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("Specify Experiment File", "No File Exception",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        private void recData (string filename)
        {
            //MessageBox.Show("In recData");
            var receiver = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);
            string receiverID = receiver.GetClientHandleAsString();

            string clientpath = @"C:\Users\Shamsuddin\Documents\ENPH 459\C# code\TestFileCreater\TestFileCreater\bin\Debug\TestFileCreater.exe";


            var startInfo = new ProcessStartInfo(clientpath,receiverID +" "+ "\""+ fileName + "\"");

            startInfo.UseShellExecute = false;
    
            
            child = Process.Start(startInfo);
            receiver.DisposeLocalCopyOfClientHandle();

            while (!child.HasExited)
            {
                using (StreamReader sr = new StreamReader(receiver))
                {
                    string line;
                    
                    try
                    {
                        if ((line = sr.ReadLine()) != null)
                        {
                            Console.WriteLine("{0}", line);
                            //MessageBox.Show(line);
                        }
                    }
                    catch (Exception e) {  }
                }
            }
            Control.CheckForIllegalCrossThreadCalls = false;
            Button2.Enabled = true;
            Control.CheckForIllegalCrossThreadCalls = true;


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Increment(+1);
        }
        //pause
        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button5.Enabled = true;
            SuspendProcess(child.Id);
        }
        //stop
        private void button4_Click(object sender, EventArgs e)
        {
            child.Kill();
            commun.Abort();
            Button2.Enabled = true;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
        }
        //resume
        private void button5_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            button3.Enabled = true;

            ResumeProcess(child.Id);
            
        }


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
