
//#define CALIBRATE
using System;
using System.Collections.Generic;
using System.Linq;
using Zaber.Serial.Core;
using OfficeOpenXml;
using System.IO.Ports;
using System.Text.RegularExpressions;
using MG17_Drivers;
using System.IO.Pipes;
using System.IO;
using System.Runtime.InteropServices;

namespace AutomaticDrugInjection
{
    class Program
    {
        // CONSTANTS DEFINED FOR LINEAR STAGES AND PUMP
        private const int MILLISEC = 1000;
        private const int WELLTOWELLSPAC_X = 4464;
        private const int WELLTOWELLSPAC_Y = 18027;
        private const int XPOS_WELL_1H = 0; 
        private const int MAX_X = 227527;
        private const int MAX_Y = 305381;
        private const int YPOS_WELL_1H = 0; 
        private const string PUMP_ADDR = "2";
        private const double ZSLIDE_OFFSET = 17.0;

        // OBJECTS REQUIRED FOR DATA STORAGE AND STAGE/PUMP AUTOMATION
        private Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();
        private ZaberBinaryPort port;
        private ZaberBinaryDevice slideX;
        private ZaberBinaryDevice slideY;
        private SerialPort pump;
        private string pump_response = "";
        //Melles Griot Variables
        private Config cf;
        private NanoSteps ns;
        private System.Array nanoStepNames;

        // Object for writing experiment log to file
        private TextWriter tw;

        // Object for writing to pipe connected to GUI
        //private StreamWriter sw;

        static void Main(string[] args)
        {
            Program program = new Program();

            string parentId = args[0];
            // create connection to pipe using id of parent process
            var sender = new AnonymousPipeClientStream(PipeDirection.Out, parentId);

            // open excel file using the file path sent in as an argument to main()
            string excelFile = args[1];
            ExcelPackage excelpack = new ExcelPackage(new System.IO.FileInfo(excelFile));

            ExcelWorksheet worksheetWithInstruction = excelpack.Workbook.Worksheets["Sheet1"];
            ExcelWorksheet worksheetWithChemToWell = excelpack.Workbook.Worksheets["Sheet2"];

#if !CALIBRATE
            string log_file = Path.GetDirectoryName(excelFile) + @"\" + Path.GetFileNameWithoutExtension(excelFile) + "_log.txt";
            program.tw = new StreamWriter(log_file, false);
            program.tw.WriteLine("Experiment Started at {0}", DateTime.Now);
#endif

            // Holds the current function (the row being executed in excel) to be executed and its parameters
            // Is emptied after each instruction is executed and then populated with the next instruction
            List<string> funcParamsList = new List<string>();
            string methodName = "";
            int row;

            // These 2 parameters are only important for the iterate/loop function. 
            // iterateRow holds the row number (in excel) from where the iteration/looping begins
            // countOfLoops holds the no. of times we have to loop over the set of instructions
            int iterateRow = 1;
            int countOfLoops = 1;

            

            // Initialization of the Zaber Slides and the Automated Syringe Pump
            program.port = new ZaberBinaryPort("COM4");
            program.pump = new SerialPort("COM3", 1200, Parity.None, 8, StopBits.One);


            //Initialization Sequence for the Melles Griot Slide
            program.cf = new Config();
            program.ns = new NanoSteps();
            program.cf.SetupAConfiguration("nanostep1");
            int numNanoSteps;
            program.cf.GetNumNanoSteps(out numNanoSteps);
            program.nanoStepNames = new string[numNanoSteps];
            program.cf.GetNanoStepNames(ref program.nanoStepNames);

            // Ensure that the Home command is done when the zaber slides are at the max position away from the needle
            // Otherwise the microplate might clash with the needle if the z-axis is homed

            using (StreamWriter sw = new StreamWriter(sender))
            {
                sw.AutoFlush = true;
                sw.WriteLine("Start");
            }
            program.ns.SingleHome(program.nanoStepNames.GetValue(0).ToString());

            program.ns.SingleMoveAbsoluteAndWait(program.nanoStepNames.GetValue(0).ToString(), ZSLIDE_OFFSET, 0);


            // Opens the ports for the zaber slides and the pump 
            program.port.Open();
            program.pump.Open();
            program.pump.DataReceived += new SerialDataReceivedEventHandler(program.DataReceivedHandler);
            // Initialize each of the linear stages
            program.slideX = new ZaberBinaryDevice(program.port, 1);
            program.slideY = new ZaberBinaryDevice(program.port, 2);
            program.slideY.Home();
            program.slideY.PollUntilIdle();

            program.slideX.MoveAbsolute(XPOS_WELL_1H);
            program.slideX.PollUntilIdle();
            program.slideY.MoveAbsolute(YPOS_WELL_1H);
            program.slideY.PollUntilIdle();

#if !CALIBRATE

            // Reads all the chemical names and their well locations and stores them in a dictionary (dict).
            for (row = 1; row <= worksheetWithChemToWell.Dimension.End.Row; row++)
            {
                program.ReadRow(worksheetWithChemToWell, row, program.dict);
            }

            // Row to start reading instructions from
            row = 9;

            // Goes row by row, reading and executing instructions sequentially
            while (row <= worksheetWithInstruction.Dimension.End.Row)
            {
                // Populates funcParamsList with the current instructions and its parameters
                program.ReadRow(worksheetWithInstruction, row, funcParamsList);

                //Get method name from first element of the list and then remove it from the list
                methodName = funcParamsList.First().Trim();
                funcParamsList.RemoveAt(0);

                // When we encounter the iterate function (from excel), we store the row we start iterating from and
                // the no. of times to iterate
                if (methodName == "Iterate")
                {
                    iterateRow = row;
                    countOfLoops = Int32.Parse(funcParamsList.First());

                }

                // If we encounter the End keyword (from excel), this is associated with the iterate function
                else if (methodName == "End")
                {
                    // If we still have to loop over the set of instruction. Change row number back to the row we started
                    // looping from
                    if (countOfLoops > 1)
                        row = iterateRow;
                    countOfLoops--;
                }

                // For all functions except iterate, we use the InvokeMethod to execute the appropriate instruction
                else
                {
                    program.InvokeMethod(methodName, funcParamsList);
                    program.tw.WriteLine("Running {0} from Row {1}: {2}", methodName, row, DateTime.Now);
                }

                // Clear the funcParamsList for the next instruction
                funcParamsList.Clear();
                row++;

            }

            program.ns.SingleMoveAbsoluteAndWait(program.nanoStepNames.GetValue(0).ToString(), ZSLIDE_OFFSET, 0);

            program.slideX.MoveAbsolute(MAX_X);
            program.slideY.Home();
            program.slideX.PollUntilIdle();
            program.slideY.PollUntilIdle();

            program.tw.WriteLine("Experiment Finished at {0}", DateTime.Now);
            program.tw.Close();
            program.port.Close();
#endif
        }


        /// <summary>
        /// Handles incoming data from KDS230 Syringe pump
        /// </summary>
        /// <param name="sender">SerialPort object for Syringe Pump</param>
        /// <param name="e">Built-in parameter</param>
        public void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            pump_response += sp.ReadExisting();
            Console.WriteLine("Data received");
        }

        /// <summary>
        /// Reads the specified row from an excel worksheet
        /// </summary>
        /// <param name="ws">The ExcelWorksheet object to read from</param>
        /// <param name="row">The row to read from</param>
        /// <param name="datatype">The object to store the data in</param>
        public void ReadRow(ExcelWorksheet ws, int row, object datatype)
        {
            // If we are reading instructions, we store it in our funcParamsList (which is of type List <string>)
            if (datatype is List<string>)
            {
                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    if (ws.Cells[row, col].Value != null)
                    {
                        ((List<string>)datatype).Add(ws.Cells[row, col].Value.ToString().Trim());
                    }
                }

            }

            // If we are reading chemical & their well locations, we store in dict 
            // (which is of type Dictionary<string, List<string>>)
            if (datatype is Dictionary<string, List<string>>)
            {
                if (ws.Cells[row, 1].Value != null)
                {
                    // Check if dictionary already contains the chemical (key) name. If not, create a new key in the dictionary
                    if (!dict.ContainsKey(ws.Cells[row, 1].Value.ToString().Trim()))
                        dict.Add(ws.Cells[row, 1].Value.ToString().Trim(), new List<string>());

                    // Read all the well locations (is in the 2nd column of the associated row) and split them into a string list
                    string parmValues = ws.Cells[row, 2].Value.ToString();
                    string[] parmVal = parmValues.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    // Store each of the well locations as a separate value for the corresponding key
                    foreach (var value in parmVal)
                    {
                        ((Dictionary<string, List<string>>)datatype)[ws.Cells[row, 1].Value.ToString().Trim()].Add(value);
                    }

                }

            }
        }

        /// <summary>
        /// Used to invoke the matching method (from excel)
        /// </summary>
        /// <param name="methodName">The name of the method to execute</param>
        /// <param name="args">The list containing the arguments for the method</param>
        public void InvokeMethod(string methodName, List<string> args)
        {
            if (args.Count != 0)
                GetType().GetMethod(methodName).Invoke(this, args.ToArray());
            else
                GetType().GetMethod(methodName).Invoke(this, null);
        }

        /// <summary>
        /// Runs the KDS 230 Pump in Withdrawal Mode
        /// </summary>
        /// <param name="flowRate">The flow rate to withdraw at</param>
        /// <param name="volWithdraw">The volume to withdraw</param>
        public void Withdraw(string flowRate, string volWithdraw)
        {
            Console.WriteLine("Infuse");

            string str;

            // regex to parse the reply data from the pump
            Regex resp_regex = new Regex(@"\r\n" + PUMP_ADDR + @"\S{1,2}");

            // example matching patterns: 1.52 ml/m, 1 ul/s, 2.5 ml, 3 ul
            Regex value_units = new Regex(@"(\d+(?:.\d+)?)(?:\s*)?([a-zA-z]+)(?:/)?([a-zA-Z]+)?");

            // rate examples: 1.52, 1
            string rate = value_units.Match(flowRate).Groups[1].ToString();
            // rate_units: mlm, uls
            string rate_units = value_units.Match(flowRate).Groups[2].ToString() + value_units.Match(flowRate).Groups[3].ToString();
            // vol: 2.5, 3
            string vol = value_units.Match(volWithdraw).Groups[1].ToString();
            // vol_units: ml, ul
            string vol_units = value_units.Match(volWithdraw).Groups[2].ToString();
            
            str = PUMP_ADDR + " mode w\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            pump_response = "";
            System.Threading.Thread.Sleep(1000);


            str = PUMP_ADDR + " ratew " + rate + " " + rate_units + "\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            pump_response = "";
            System.Threading.Thread.Sleep(1000);

            str = PUMP_ADDR + " volw " + vol + " " + vol_units + "\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            pump_response = "";
            System.Threading.Thread.Sleep(1000);

            str = PUMP_ADDR + " run\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            System.Threading.Thread.Sleep(1000);

            while (pump_response[pump_response.Length - 1] != ':')
            {
                pump_response = "";
                System.Threading.Thread.Sleep(5000);
                str = PUMP_ADDR + " run?\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                Console.WriteLine("{0}", "Inside Loop");
                continue;
            }
            pump_response = "";

        }

        /// <summary>
        /// Runs the KDS 230 Pump in Infusion Mode
        /// </summary>
        /// <param name="flowRate">The flow rate to infuse at</param>
        /// <param name="volInfused">The volume to infuse</param>
        public void Infuse(string flowRate, string volInfused)
        {
            Console.WriteLine("Infuse");

            string str;

            // regex to parse the reply data from the pump
            Regex resp_regex = new Regex(@"\r\n" + PUMP_ADDR + @"\S{1,2}");

            // example matching patterns: 1.52 ml/m, 1 ul/s
            Regex value_units = new Regex(@"(\d+(?:.\d+)?)(?:\s*)?([a-zA-z]+)(?:/)?([a-zA-Z]+)?");

            string rate = value_units.Match(flowRate).Groups[1].ToString();
            string rate_units = value_units.Match(flowRate).Groups[2].ToString() + value_units.Match(flowRate).Groups[3].ToString();
            string vol = value_units.Match(volInfused).Groups[1].ToString();
            string vol_units = value_units.Match(volInfused).Groups[2].ToString();

            str = PUMP_ADDR + " mode i\r";
            Console.WriteLine("{0}",str);
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            pump_response = "";
            System.Threading.Thread.Sleep(1000);


            str = PUMP_ADDR + " ratei " + rate + " " + rate_units + "\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            pump_response = "";
            System.Threading.Thread.Sleep(1000);

            str = PUMP_ADDR + " voli " + vol + " " + vol_units + "\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            pump_response = "";
            System.Threading.Thread.Sleep(1000);

            str = PUMP_ADDR + " run\r";
            pump.Write(str);
            while (!resp_regex.Match(pump_response).Success)
                continue;
            Console.WriteLine("{0}", pump_response);
            System.Threading.Thread.Sleep(1000);

            while (pump_response[pump_response.Length - 1] != ':')
            {
                pump_response = "";
                System.Threading.Thread.Sleep(5000);
                str = PUMP_ADDR + " run?\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                continue;
            }
            pump_response = "";

        }

        /// <summary>
        /// Runs the linear stages to position the needle inside the required well
        /// </summary>
        /// <param name="chemical">The name of the chemical to move to</param>
        public void MoveTo(string chemical)
        {
            //We should first lower the melles griot slide to get out of the current well
            double z_pos;
            ns.SingleGetPosition(nanoStepNames.GetValue(0).ToString(), out z_pos);
            if (!(Math.Abs(z_pos - ZSLIDE_OFFSET) < 0.1))
            {
                ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), ZSLIDE_OFFSET, 0);
            }

            Console.WriteLine("move");
            // Initialize a list to store all the well locations specified in the dictionary
            var regex_wellCol = new Regex(@"\d+");
            var regex_wellRow = new Regex(@"(?i)[A-H]{1}");
            List<string> wellNums = dict[chemical];
            int nextCol = Int32.Parse((regex_wellCol.Match(wellNums.First())).ToString());
            int nextRow = regex_wellRow.Match(wellNums.First()).ToString()[0] - 'A';

            // Specifies the absolute positions of the well to move to
            int nextPosX = (7 - nextRow) * WELLTOWELLSPAC_X;
            int nextPosY = (nextCol - 1) * WELLTOWELLSPAC_Y;

            slideX.MoveAbsolute(nextPosX);
            slideY.MoveAbsolute(nextPosY);

            slideX.PollUntilIdle();
            slideY.PollUntilIdle();

            // Move the microplate well to the needle
            ns.SingleHome(nanoStepNames.GetValue(0).ToString());

            Console.WriteLine("WellColNum: {0}", nextCol);
            Console.WriteLine("WellRowNum: {0}", nextRow);
            Console.WriteLine("WellNo: {0}", wellNums.First());

            // ASSUMING ALL CHEMICAL IS USED UP AFTER ONE OPERATION (MIGHT NEED TO CHANGE)
            dict[chemical].Remove(wellNums.First());

        }
        /// <summary>
        /// The method pauses the program before executing the next instruction
        /// </summary>
        /// <param name="seconds">Time (in s) to pause for</param>
        public void Pause(string seconds)
        {
            Console.WriteLine("Pause");
            System.Threading.Thread.Sleep(Int32.Parse(seconds) * MILLISEC);
        }

        /// <summary>
        /// Executes a 3 step Rinse cycle going to wells 1H, 2H and 3H
        /// </summary>
        /// <param name="time">The time to pause for in each well</param>
        public void Rinse(string time)
        {
            double zpos;
            ns.SingleGetPosition(nanoStepNames.GetValue(0).ToString(), out zpos);
            if (!(Math.Abs(zpos - ZSLIDE_OFFSET) < 0.1))
            {
                ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), ZSLIDE_OFFSET, 0);
            }

            slideX.MoveAbsolute(XPOS_WELL_1H);
            slideX.PollUntilIdle();
            for (int count = 0; count < 3; count++)
            {
                if (count != 0)
                    ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), ZSLIDE_OFFSET, 0);
                slideY.MoveAbsolute(WELLTOWELLSPAC_Y * count);
                slideY.PollUntilIdle();
                ns.SingleHome(nanoStepNames.GetValue(0).ToString());
                System.Threading.Thread.Sleep(int.Parse(time)*1000);
            }

            ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), ZSLIDE_OFFSET, 0);
        }

    }
}

