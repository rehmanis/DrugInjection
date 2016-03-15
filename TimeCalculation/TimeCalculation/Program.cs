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
        // rinsing time in seconds
        private const double RINSE_TIME = 1.0;

        private const double MILLI = 1e-3;

        // All speed are in m/s
        private const int X_SPEED = 1; 
        private const int Y_SPEED = 1;
        private const int Z_SPEED = 1;

        // First well row and col location of where rinsing of the needle begins
        private const int WELL_X_REF = 7;
        private const int WELL_Y_REF = 1;

        // well to well distance is 9 mm
        private const int WELL_DIS = 9;

        // Total time for running the one experiment
        private double totalTime = 0.0;
        private string lastChem = "";
        
        // intialize a mulit value dictionary to store chemcial as key and their well locations as list of values
        private Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();
        private int curColPos = 1;
        private int curRowPos = 7;

        static void Main(string[] args)
        {
        
            Program program = new Program();
            string parentSenderId;

            parentSenderId = args[0];

            var sender = new AnonymousPipeClientStream(PipeDirection.Out, parentSenderId);

            // get the excel file path from the GUI 
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
                    program.AddTime(methodName, funcParamsList);

                // Clear the funcParamsList for the next instruction
                funcParamsList.Clear();
                row++;

            }

            using (StreamWriter sw = new StreamWriter(sender))
            {
                Console.WriteLine("using streamwriter");
                sw.AutoFlush = true;
                sw.WriteLine("{0}", program.totalTime.ToString());
            }
            Console.ReadLine();
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

        public void AddTime(string method, List<string> funcParamsList )
        {




            switch (method)
            {
                case "Pause":
                    Console.WriteLine("Pause");
                    // add the time the needle pauses at a well to the total time for running the experiment.
                    // This is stored in the FuncParamsList
                    totalTime += Int32.Parse(funcParamsList.First());
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
                        totalTime += CalculateTime(vol, flowrate, rateTimeUnits, 1);
                    }
                    // flow rate has a smaller volume unit than vol's volume unit. Example, flowrate is ul/s and vol is ml
                    // for this experiment, we are assuming that only ml and ul are the only options
                    else if (result > 0)
                    {
                        Console.WriteLine("flowrate volume units is less than vol volume unit");
                        totalTime += CalculateTime(vol, flowrate,rateTimeUnits, 1/MILLI);

                    }
                    // flow rate has a larger volume unit than vol's volume unit.
                    else
                    {
                        Console.WriteLine("flowrate volume units is greater than vol volume unit");
                        totalTime += CalculateTime(vol, flowrate,rateTimeUnits, MILLI);
                    }

                    break;

                case "MoveTo":

                    Console.WriteLine("MOVE");

                    string currentChem = funcParamsList[0];
                    Console.WriteLine(currentChem);

                    var regex_wellCol = new Regex(@"\d+");
                    var regex_wellRow = new Regex(@"(?i)[A-H]{1}");
                   

                    List<string> wellNums = dict[currentChem];
                    foreach(string wellN in wellNums)
                    {
                        Console.WriteLine(wellN);
                    }

                    int nextColPos = int.Parse((regex_wellCol.Match(wellNums.First())).ToString()); // used for moving yslide
                    int nextRowPos = regex_wellRow.Match(wellNums.First()).ToString()[0] - 'A'; // used for moving xslide

            
                    string nextChemConc = (regex_wellCol.Match(currentChem)).ToString();
                    string lastChemConc = (regex_wellCol.Match(lastChem)).ToString();

                    // The only reason to avoid a rinse cycle, will be if the last chemical and the current chemical
                    // are the same and the current one is at a higher concentration than the last
                    if (nextChemConc == "" || lastChemConc == "" || !(Double.Parse(nextChemConc) >
                        Double.Parse(lastChemConc) && String.Equals(currentChem, lastChem, StringComparison.Ordinal)))
                    {
                        // if rinsing cycle is needed than we have to add the time it takes for the needle to traverse to
                        // well H1 where the water for rinsing is kept. A rinsing time is also added to the total time.
                        totalTime +=  WELL_DIS*(Math.Abs(curColPos-WELL_Y_REF)/Y_SPEED + 
                            Math.Abs(curRowPos - WELL_X_REF) / X_SPEED) + RINSE_TIME;
                    }

                    // time it takes for needle to traverse to next well location is added.
                    totalTime += WELL_DIS * (Math.Abs(curRowPos - nextRowPos) / X_SPEED + Math.Abs(curColPos - nextColPos)/Y_SPEED);

                    Console.WriteLine("WellColNum: {0}", nextColPos);
                    Console.WriteLine("WellRowNum: {0}", nextRowPos);
                    Console.WriteLine("WellNo: {0}", wellNums.First());

                    // ASSUMING ALL CHEMICAL IS USED UP AFTER ONE OPERATION (MIGHT NEED TO CHANGE)
                    dict[currentChem].Remove(wellNums.First());

                    // current row and column position is updated to next ones.
                    curRowPos = nextRowPos;
                    curColPos = nextColPos;
                    lastChem = currentChem;
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

            if (rateTUnits.Equals("s"))
            {
                Console.WriteLine("time units: s");
                time = convFactor * (volume / rate);
            }
                
            else if (rateTUnits.Equals("min"))
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
