using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO.Pipes;
using System.IO;

namespace TimeCalculation
{
    class Program
    {
        private const double MILLI = 1e-3;

        // All speed are in mm/s
        private const double X_SPEED = 54.5; 
        private const double Y_SPEED = 13.6;

        private const double Z_OFFSET = 17.0; //in mm
        // Time for Z-Slide
        private const double Z_HOME_TIME = Z_OFFSET;
        private const double Z_DOWN_TIME = (9.5 / 17.0) * Z_OFFSET;

        // well to well distance is 9 mm
        private const int WELL_DIS = 9;

        // Total time for running the one experiment
        private double totalTime = 0.0;
        private int curExlRow = 0;

        int curRowPos = 7, curColPos = 1, nextRowPos, nextColPos;
        bool in_well = false;

        static void Main(string[] args)
        {
        
            Program program = new Program();
            
            string parentSenderId;

            parentSenderId = args[0];

            var sender = new AnonymousPipeClientStream(PipeDirection.Out, parentSenderId);
            StreamWriter sw = new StreamWriter(sender);
            sw.AutoFlush = true;

            // get the excel file path from the GUI 
            string excelFile = args[1]; //@"C:\Documents and Settings\Admin\Desktop\automatedDrugTest.xlsx";
            ExcelPackage excelpack = new ExcelPackage(new System.IO.FileInfo(excelFile));

            ExcelWorksheet worksheetWithInstruction = excelpack.Workbook.Worksheets["Sheet1"];

            // Before starting the experiment, the z-slide home and goes back down, and the x-y slide go to their home positions
            program.totalTime += Z_HOME_TIME + Z_DOWN_TIME;
            program.totalTime += 450 / X_SPEED;

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

            // Row to start reading instructions from
            row = 9;

            // Goes row by row, reading and executing instructions sequentially
            while (row <= worksheetWithInstruction.Dimension.End.Row)
            {
                program.curExlRow = row;
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
                    program.AddTime(methodName, funcParamsList);

                // Clear the funcParamsList for the next instruction
                funcParamsList.Clear();
                row++;

            }


            program.totalTime += Z_DOWN_TIME;

            sw.WriteLine("{0}", program.totalTime.ToString());

            Console.WriteLine("{0}", program.totalTime);
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
        }

        public void AddTime(string method, List<string> funcParamsList )
        {
            double z_move_time;
            int pause_time;
            double inf_wrw_time;
            switch (method)
            {
                case "Pause":
                    Console.WriteLine("Pause");
                    // add the time the needle pauses at a well to the total time for running the experiment.
                    // This is stored in the FuncParamsList
                    pause_time = Int32.Parse(funcParamsList.First());
                    totalTime += pause_time;
                    break;

                case "Rinse":
                    Console.WriteLine("Rinse");
                    nextRowPos = 7;
                    nextColPos = 1;
                    
                    if (!in_well)
                        z_move_time = 3 * (Z_HOME_TIME + Z_DOWN_TIME);
                    else
                        z_move_time = 3 * Z_HOME_TIME + 4 * Z_DOWN_TIME;

                    double xy_move_time = WELL_DIS * (Math.Abs(curRowPos - nextRowPos) / X_SPEED + Math.Abs(curColPos - nextColPos) / Y_SPEED)
                         + 2 * WELL_DIS * (1 / Y_SPEED);
                    double rinse_time = 3 * int.Parse(funcParamsList.First());
                    totalTime += rinse_time + xy_move_time + z_move_time;
                    curRowPos = 7;
                    curColPos = 3;
                    in_well = false;
                    break;

                case "Infuse": case "Withdraw":
                    Console.WriteLine("Withdraw or Infuse");

                    // intialize the regex pattern for extracting flowrate and volume values and units
                    string pattern = @"(\d+(?:.\d+)?)(?:\s*)?([a-zA-Z]+)(?:/)?([a-zA-Z]+)?";

                   // Match matchVol = Regex.Match("5 ml", pattern);
                    Match matchVol = Regex.Match(funcParamsList[1].Trim(), pattern);
                    Match matchRate = Regex.Match(funcParamsList[0], pattern);

                    Console.WriteLine(funcParamsList[1]);
                    Console.WriteLine(funcParamsList[0]);

                    double vol = double.Parse((matchVol.Groups[1]).ToString());
                    double flowrate = double.Parse((matchRate.Groups[1]).ToString());


                    string volUnits = (matchVol.Groups[2]).ToString();
                    string rateVolUnits = matchRate.Groups[2].ToString();
                    string rateTimeUnits = matchRate.Groups[3].ToString();

                   
                    Console.WriteLine("{0}", matchVol.Length);
                    Console.WriteLine("{0}", matchRate.Length);
                    Console.WriteLine("volume: {0}",vol);
                    Console.WriteLine("volume unit: {0}", volUnits);
                    Console.WriteLine("flowrate: {0}",flowrate);
                    Console.WriteLine("flowrate volume unit: {0}", rateVolUnits);
                    Console.WriteLine("flowrate time unit: {0}", rateTimeUnits);

                    int result = string.CompareOrdinal(rateVolUnits, volUnits);
                   
                    //flowrate and vol has same volume unit. Example flowrate is ml/s and vol is ml
                    if (result == 0)
                    {
                        Console.WriteLine("volume units is same");
                        inf_wrw_time = CalculateTime(vol, flowrate, rateTimeUnits, 1);
                    }
                    // flow rate has a smaller volume unit than vol's volume unit. Example, flowrate is ul/s and vol is ml
                    // for this experiment, we are assuming that only ml and ul are the only options
                    else if (result > 0)
                    {
                        Console.WriteLine("flowrate volume units is less than vol volume unit");
                        inf_wrw_time = CalculateTime(vol, flowrate,rateTimeUnits, 1/MILLI);

                    }
                    // flow rate has a larger volume unit than vol's volume unit.
                    else
                    {
                        Console.WriteLine("flowrate volume units is greater than vol volume unit");
                        inf_wrw_time = CalculateTime(vol, flowrate,rateTimeUnits, MILLI);
                    }

                    pause_time = 4;
                    totalTime += (pause_time + inf_wrw_time);
                    
                    break;

                case "MoveTo":

                    Console.WriteLine("MOVE");
                    string currentChem = funcParamsList[0];
                    Console.WriteLine(currentChem);

                    var regex_wellCol = new Regex(@"\d{1,2}");
                    var regex_wellRow = new Regex(@"[A-Ha-h]{1}");
                    var regex_row_col = new Regex(@"\(\s*([A-Ha-h0-9]+)\s*\)");


                    string well = regex_row_col.Match(currentChem).Groups[1].ToString();
                    nextColPos = Int32.Parse(regex_wellCol.Match(well).ToString());
                    nextRowPos = regex_wellRow.Match(well).ToString()[0] - 'A';
                    // time it takes for needle to traverse to next well location is added.

                    if (!in_well)
                        z_move_time = Z_HOME_TIME;
                    else
                        z_move_time = Z_HOME_TIME + Z_DOWN_TIME;

                    totalTime += WELL_DIS * (Math.Abs(curRowPos - nextRowPos) / X_SPEED + Math.Abs(curColPos - nextColPos)/Y_SPEED) + 
                        z_move_time;



                    // current row and column position is updated to next ones.
                    curRowPos = nextRowPos;
                    curColPos = nextColPos;
                    in_well = true;
                    break;

                    
            }


        }
        /// <summary>
        /// return a time calculation after making the units of flow rate and volume same
        /// </summary>
        /// <param name="volume"> volume infused or withdrawn from the MILLIwell plate</param>
        /// <param name="rate"> the flow rate at which the volume is infused or withdrawn </param>
        /// <param name="convFactor"> the conversion factor needed to make the volume units same</param>
        /// <returns></returns>
        public double CalculateTime (double volume, double rate, string rateTUnits, double convFactor)
        {
            double time = 0.0;
            Console.WriteLine("Time component of rate is: {0}", rateTUnits);
            if (rateTUnits.Equals("s", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("time units: s");
                time = convFactor * (volume / rate);
            }

            else if (rateTUnits.Equals("m", StringComparison.OrdinalIgnoreCase))
            {
                time = convFactor * 60 * (volume / rate);
                Console.WriteLine("time units: min");
            }
                
            else
            {
                time = convFactor * 3600 * (volume / rate);
                Console.WriteLine("time units: hr");
            }
                

            Console.WriteLine("inj/with: {0}", time);

            return time;
        }
    }

}
