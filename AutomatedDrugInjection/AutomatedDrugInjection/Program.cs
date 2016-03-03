
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

namespace AutomaticDrugInjection
{
    class Program
    {
        private const int MILLISEC = 1000;
        private const int WELLTOWELLSPAC_X = (int)(8.85e-3 / 1.984375e-6);
        private const int WELLTOWELLSPAC_Y = (int)(9e-3 / 0.49609375e-6);
        private const int XPOS_WELL_1H = 0; //(int)(140e-3 / 1.984375e-6);
        private const int MAX_X = 227527;
        private const int MAX_Y = 305381;
        private const int YPOS_WELL_1H = 0; //(int)(140e-3 / 0.49609375e-6);
        private const string PUMP_ADDR = "2";
       
        private const double SPEEDX = 1; // speed of x-slide in mm/s
        private const double SPEEDY = 1;  // speed of y-slide in mm/s
        private const double SPEEDZ = 1; // speed of z-slide in mm/s

        private string lastChem = "";
        private Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();
        private ZaberBinaryPort port;
        private ZaberBinaryDevice slideX;
        private ZaberBinaryDevice slideY;
        private SerialPort pump;
        private string pump_response = "";
        private int totalTime = 0;
        private bool TIME_CALC = false;

        //private Program programObj = new Program();

        //Melles Griot Variables
        private Config cf;
        private NanoSteps ns;
        private System.Array nanoStepNames;

        static void Main(string[] args)
        {
            string parentSenderId;
            Program program = new Program();


            parentSenderId = args[0];
            Console.WriteLine("{0}", args[0]);
            Console.WriteLine("{0}", args[1]);

            var sender = new AnonymousPipeClientStream(PipeDirection.Out, parentSenderId);

            Console.WriteLine("outside stream");

            if (args.Length == 3)
            {
                program.TIME_CALC = true;

            }
       


            string excelFile = args[1]; //@"C:\Documents and Settings\Admin\Desktop\automatedDrugTest.xlsx";
            ExcelPackage excelpack = new ExcelPackage(new System.IO.FileInfo(excelFile));

            ExcelWorksheet worksheetWithInstruction = excelpack.Workbook.Worksheets["Sheet1"];
            ExcelWorksheet worksheetWithChemToWell = excelpack.Workbook.Worksheets["Sheet2"];

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

            

            if(program.TIME_CALC == false)
            {
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
                System.Array velParams;
                velParams = new Double[3];
                velParams.SetValue(1, 0);
                velParams.SetValue(2, 1);
                velParams.SetValue(6, 2);
                program.ns.SingleSetVelocityProfile(program.nanoStepNames.GetValue(0).ToString(), ref velParams);
                // Ensure that the Home command is done when the zaber slides are at the max position away from the needle
                // Otherwise the microplate might clash with the needle if the z-axis is homed
                program.ns.SingleHome(program.nanoStepNames.GetValue(0).ToString());
                program.ns.SingleMoveAbsoluteAndWait(program.nanoStepNames.GetValue(0).ToString(), 5.0, 0);
                program.ns.SingleMoveAbsoluteAndWait(program.nanoStepNames.GetValue(0).ToString(), 10.0, 0);


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
            }
            

            // Reads all the chemical names and their well locations and stores them in a dictionary (dict).
            for (row = 1; row <= worksheetWithChemToWell.Dimension.End.Row; row++)
            {
                program.ReadRow(worksheetWithChemToWell, row, program.dict);
            }

            // Row to start reading instructions from
            row = 1;

            // Goes row by row, reading and executing instructions sequentially
            while (row <= worksheetWithInstruction.Dimension.End.Row)
            {
                // Populates funcParamsList with the current instructions and its parameters
                program.ReadRow(worksheetWithInstruction, row, funcParamsList);

                //Get the method name from first element of the list, ignoring the white spaces, and later
                //remove it from the list.
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
                    program.InvokeMethod(methodName, funcParamsList);

                // Clear the funcParamsList for the next instruction
                funcParamsList.Clear();
                row++;

            }

            if (program.TIME_CALC == false)
            {
                program.ns.SingleMoveAbsoluteAndWait(program.nanoStepNames.GetValue(0).ToString(), 5.0, 0);
                program.ns.SingleMoveAbsoluteAndWait(program.nanoStepNames.GetValue(0).ToString(), 10.0, 0);

                program.slideX.MoveAbsolute(MAX_X);
                program.slideY.Home();
                program.slideX.PollUntilIdle();
                program.slideY.PollUntilIdle();

                program.port.Close();
            }
            else
            {
                using (StreamWriter sw = new StreamWriter(sender))
                {
                    Console.WriteLine("using streamwriter");
                    sw.AutoFlush = true;
                    sw.WriteLine("{0}", program.totalTime.ToString());
                }
            }



        }

        /// <summary>
        /// Is meant to handle data coming in from the serial port for the Melles Griot linear stage
        /// </summary>
        /// <param name="sender">Is the SerialPort object associated with the SerialPort</param>
        /// <param name="e">Built-in parameter</param>
        public void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            pump_response += sp.ReadExisting();
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
            Console.WriteLine("Withdraw");

            if (TIME_CALC == false)
            {
                string str;
                string rate_units = @"mlm";
                string vol_units = @"ml";
                Regex resp_regex = new Regex(@"\r\n" + PUMP_ADDR + @"\S{1,2}");

                str = PUMP_ADDR + " mode w\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                pump_response = "";
                System.Threading.Thread.Sleep(2000);


                str = PUMP_ADDR + " ratew " + flowRate + " " + rate_units + "\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                pump_response = "";
                System.Threading.Thread.Sleep(2000);

                str = PUMP_ADDR + " volw " + volWithdraw + " " + vol_units + "\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                pump_response = "";
                System.Threading.Thread.Sleep(2000);

                str = PUMP_ADDR + " run\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                System.Threading.Thread.Sleep(2000);

                while (pump_response[pump_response.Length - 1] != ':')
                {
                    pump_response = "";
                    System.Threading.Thread.Sleep(10000);
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
            else 
                totalTime += (int)(Double.Parse(volWithdraw) / Double.Parse(flowRate));


        }

        /// <summary>
        /// Runs the KDS 230 Pump in Infusion Mode
        /// </summary>
        /// <param name="flowRate">The flow rate to infuse at</param>
        /// <param name="volInfused">The volume to infuse</param>
        public void Infuse(string flowRate, string volInfused)
        {
            Console.WriteLine("Infuse");

            if (TIME_CALC == false)
            {
                string str;
                string rate_units = @"mlm";
                string vol_units = @"ml";
                Regex resp_regex = new Regex(@"\r\n" + PUMP_ADDR + @"\S{1,2}");

                str = PUMP_ADDR + " mode i\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                pump_response = "";
                System.Threading.Thread.Sleep(2000);


                str = PUMP_ADDR + " ratei " + flowRate + " " + rate_units + "\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                pump_response = "";
                System.Threading.Thread.Sleep(2000);

                str = PUMP_ADDR + " voli " + volInfused + " " + vol_units + "\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                pump_response = "";
                System.Threading.Thread.Sleep(2000);

                str = PUMP_ADDR + " run\r";
                pump.Write(str);
                while (!resp_regex.Match(pump_response).Success)
                    continue;
                Console.WriteLine("{0}", pump_response);
                System.Threading.Thread.Sleep(2000);

                while (pump_response[pump_response.Length - 1] != ':')
                {
                    pump_response = "";
                    System.Threading.Thread.Sleep(10000);
                    str = PUMP_ADDR + " run?\r";
                    pump.Write(str);
                    while (!resp_regex.Match(pump_response).Success)
                        continue;
                    Console.WriteLine("{0}", pump_response);
                    pump_response = "";
                    continue;
                }

            }
            else
                totalTime += (int)(Double.Parse(volInfused) / Double.Parse(flowRate));

        }

        /// <summary>
        /// Runs the linear stages to position the needle inside the required well
        /// </summary>
        /// <param name="chemical">The name of the chemical to move to</param>
        public void MoveTo(string chemical)
        {
            if (TIME_CALC == false)
            {
                //We should first lower the melles griot slide to get out of the current well
                double z_pos;
                ns.SingleGetPosition(nanoStepNames.GetValue(0).ToString(), out z_pos);
                if (!(Math.Abs(z_pos - 10.0) < 0.1))
                {
                    ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 5.0, 0);
                    ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 10.0, 0);
                }
            }


            Console.WriteLine("move");
            // Initialize a list to store all the well locations specified in the dictionary
            List<string> wellNums = dict[chemical];
            int nextCol = (int)char.GetNumericValue(((wellNums.First())[0]));
            int nextRow = ((wellNums.First())[1]) - 'A';

            // Regular expression to define the syntax for concentration (e.g. 3.5M)
            // If it is 3.5 M, this will extract 3.5 (MIGHT NEED TO RECONSIDER THIS REGULAR EXPRESSION)
            var regex = new Regex(@"\d+(.\d+)?");
            string nextChemConc = (regex.Match(chemical)).ToString();
            string lastChemConc = (regex.Match(lastChem)).ToString();


            // Specifies the absolute positions of the well to move to
            int nextPosX = (7 - nextRow) * WELLTOWELLSPAC_X;
            int nextPosY = (nextCol - 1) * WELLTOWELLSPAC_Y;

            // The only reason to avoid a rinse cycle, will be if the last chemical and the current chemical
            // are the same and the current one is at a higher concentration than the last
            if (nextChemConc == "" || lastChemConc == "" || !(Double.Parse(nextChemConc) > Double.Parse(lastChemConc)
                && String.Equals(chemical, lastChem, StringComparison.Ordinal)))
            {
                if (TIME_CALC == false)
                {
                    Rinse();
                }
                
                totalTime += (int)(6*10/SPEEDZ + 3 * 10/SPEEDX) ;
            }
                


            if (TIME_CALC == false)
            {
                slideX.MoveAbsolute(nextPosX);
                slideY.MoveAbsolute(nextPosY);

                slideX.PollUntilIdle();
                slideY.PollUntilIdle();

                // Move the microplate well to the needle
                ns.SingleHome(nanoStepNames.GetValue(0).ToString());
            }
            else
                //totalTime += (int)((Math.Abs(nextPosX))/SPEEDX + 
                    //Math.Abs(nextPosX)/SPEEDY);

            Console.WriteLine("WellColNum: {0}", nextCol);
            Console.WriteLine("WellRowNum: {0}", nextRow);
            Console.WriteLine("WellNo: {0}", wellNums.First());

            // ASSUMING ALL CHEMICAL IS USED UP AFTER ONE OPERATION (MIGHT NEED TO CHANGE)
            dict[chemical].Remove(wellNums.First());
            // update last chemical
            lastChem = chemical;
            Console.WriteLine("done move");
        }
        /// <summary>
        /// The method pauses the program before executing the next instruction
        /// </summary>
        /// <param name="seconds">Time (in s) to pause for</param>
        public void Pause(string seconds)
        {
            Console.WriteLine("Pause");
            System.Threading.Thread.Sleep(Int32.Parse(seconds) * MILLISEC);
            totalTime += Int32.Parse(seconds);
        }

        /// <summary>
        /// Completes a rinse cycle between moving from one chemical to another
        /// </summary>
        public void Rinse()
        {
            slideX.MoveAbsolute(XPOS_WELL_1H);
            slideX.PollUntilIdle();
            for (int count = 0; count < 3; count++)
            {
                if (count == 0)
                {
                    double zpos;
                    ns.SingleGetPosition(nanoStepNames.GetValue(0).ToString(), out zpos);
                    if (!(Math.Abs(zpos - 10.0) < 0.1))
                    {
                        ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 5.0, 0);
                        ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 10.0, 0);
                    }
                }
                else
                {
                    ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 5.0, 0);
                    ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 10.0, 0);
                }
                slideY.MoveAbsolute(WELLTOWELLSPAC_Y * count);
                slideY.PollUntilIdle();
                ns.SingleHome(nanoStepNames.GetValue(0).ToString());
                System.Threading.Thread.Sleep(1000);
            }

            ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 5.0, 0);
            ns.SingleMoveAbsoluteAndWait(nanoStepNames.GetValue(0).ToString(), 10.0, 0);
        }


    }
}

