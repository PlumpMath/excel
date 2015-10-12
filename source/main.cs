/*
 *  Program to allow Excel to be used in batch automatization
 *  
 *  The program opens an excel file in the background,
 *  pastes content from a file into cells,
 *  runs macros, then saves charts and outputs selected cells
 *  
 *  This gives a lot of posibilities:
 *      - You only need to know excel to create batch programs
 *      - Data can easily be checked by outputting charts of relevant data. Visual inspection of the charts is easy and quick.
 *      -
 *      
 *  
 *  martin.svendsen@akersolutions, 2014
 *  
 */
 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using AkerSolutions;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Globalization;

namespace ExcelMacro
{

    class Program
    {

        // create shared objects for the excel file
        // this is important to easily dispose of the excel file if something goes wrong,
        // since the file runs in a background prosess, and not as a subprocess
        static Excel.Application oXL = null;
        static Excel._Workbook oWB = null;
        static Excel._Worksheet oSheet;

        static void Main(string[] args)
        {
            // set english culture (for english function names and . decimal)
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-GB");

            // list seperator is now , Trying to change it to ; doesnt work :(
            // System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator = ";";

            // show manual
            bool showMan = false;

            // check pipe
            String pipedText = "";
            bool isKeyAvailable;
            bool piped = false;
            try { isKeyAvailable = System.Console.KeyAvailable; }
            catch { pipedText = System.Console.In.ReadToEnd(); piped = true; }

            // if no args or pipe, show manual
            if (piped == false && args.Length == 0) showMan = true;

            // set default values
            string infile = "";
            string paste = "";
            List<string> macro = new List<string>();
            int[] cellA = new int[2] { 0, 1 };
            int[] cellB = new int[2] { 0, 1 };
            int[] outCellA = new int[2] { 0, 1 };
            int[] outCellB = new int[2] { 0, 0 };
            string sheet = "";
            string active = "";

            // warnings are off by default, since excel will warn about ANYTHING, which interupts the program and leads to errors.
            // f.eks. log charts will make the program fail, when they are given blank data in the step before new data is provided.
            bool warning = false;

            // save on exit
            bool save = true;

            // convert error codes to text in output
            bool outErr = true;

            // paste as text
            bool asText = false;


            // blehh..
            string errLine = "--------------------------------------------------------------------------------";

            // tab as default output space character
            string cellSpacer = "\t";


            // all charts that should be saved
            List<string> charts = new List<string>();


            // check input arguments
            int iarg = 0;
            for (int i = 0; i < args.Length;i++ )
            {
                if (args[i].StartsWith("-")) {

                    // show manual
                    if (args[i] == "-help" || args[i] == "--help" || args[i] == "-?") showMan = true;

                    // specify macro to run
                    if (args[i] == "-m") {
                      try {
                          macro.Add(args[i + 1]);
                        i++;
                      }
                      catch {
                        Error("No macro name given for -m.",1);
                      }  
                    }
                    // paste input as text?
                    if (args[i] == "-t") asText = true;
                    // dont save
                    if (args[i] == "-n") save = false;
                    // hide warnings
                    if (args[i] == "-w") warning = true;
                    // set space character
                    if (args[i] == "-b")
                    {
                        cellSpacer = " ";
                        if (args.Length > i+1)
                            if (args[i + 1].Length == 1)
                            {
                                cellSpacer = args[i + 1];
                                i++;
                            }
                    }
                    // set paste sheet
                    if (args[i] == "-p")
                    {
                        try
                        {
                            active = args[i + 1];
                            i++;
                        }
                        catch
                        {
                            Error("No paste name given for -p.", 1);
                        }
                    } 
                    // set output sheet
                    if (args[i] == "-s")
                    {
                        try
                        {
                            sheet = args[i + 1];
                            i++;
                        }
                        catch
                        {
                            Error("No sheet name given for -s.", 1);
                        }
                    }
                    // blank errors
                    if (args[i] == "-#") outErr = false;
                }

                else {
                    // excel file
                    if (iarg == 0) infile = args[i];
                    // paste file
                    else if (iarg == 1 && !piped)
                    {
                        paste = args[i];
                        if (paste == "~" || paste == "")
                        {
                            paste = "";
                            iarg++; iarg++;
                        }
                    }
                    // input cell ref
                    else if (iarg == 2)
                    {
                        string[] cellArr = args[i].Split(':');
                        if (cellArr.Length == 1) {
                             cellA = ExcelCellRef(cellArr[0]);
                        }
                        else {
                            cellA = ExcelCellRef(cellArr[0]);
                            cellB = ExcelCellRef(cellArr[1]);
                            iarg++;
                        }
                    }

                    else if (iarg == 3)
                    {
                        cellB = ExcelCellRef(args[i]);
                    }
                    // output cell ref
                    else if (iarg == 4)
                    {
                        string[] cellArr = args[i].Split(':');
                        if (cellArr.Length == 1)
                        {
                           outCellA = ExcelCellRef(cellArr[0]);
                        }
                        else
                        {
                            outCellA = ExcelCellRef(cellArr[0]);
                            outCellB = ExcelCellRef(cellArr[1]);
                            iarg++;
                        }
                    }
                    else if (iarg == 5)
                    {
                        outCellB = ExcelCellRef(args[i]);
                    }

                    // output charts
                    else if (iarg > 5)
                    {
                        charts.Add(args[i]);
                    }

                    iarg++;
                }
            }

            // Print header
            if (showMan)
            {
                Print(@"Usage: excel [OPTIONS] ExcelFile PasteFile Cell1 Cell2 OutCell1 OutCell2 [Chart1 [Chart2 ..]]
   or: excel =FORMULA

Opens ExcelFile and places the contents of PasteFile from position given.
Position is the range between Cell1 and Cell2. Unused cells are cleared.
Echos all filled rows in the range between OutCell1 and OutCell2.
Any charts named will be saved to <ExcelFile_ChartN>.png.
'Sheet.ChartN' can be used if multiple charts has the same name.

    -p PasteSheet   Select the sheet that should be pasted to.
    -s OutSheet     Select the sheet that should be outputted.
    -m Macro        Run macro after paste. If -m is used multiples
                    times, more than 1 macro can be executed.
    -n              Do not save workbook
    -w              display Excel dialogs. Default is to surpress.
    -#              Replace errors with blanks in output
    -b [CHAR]       Set cell-spacing character in output to 
                    'space' or 'CHAR' (default is 'tab')
    -t              Insert PasteFile as text instead of numbers

If '~' is specified as PasteFile no file is loaded and Cell1 and Cell2 
should not be specified.

Version 1.0. Report bugs to <martsve@gmail.com>");
                Environment.Exit(0);
            }



            if (infile.StartsWith("="))
            {
                string result = "";
                try
                {
                    result = ExcelMath.Calc(infile);
                }
                catch (Exception ex) {

                    Console.Error.WriteLine("Error: " + ex.Message);
#if DEBUG
                    Console.ReadKey();
#endif
                    Environment.Exit(1);
                }

                Print(result);
#if DEBUG
                Console.ReadKey();
#endif
                Environment.Exit(0);
            }

            // open file
            if (piped == false && infile == "") Error("No file given.", 1);
            System.IO.TextReader stream = new StringReader(pipedText);
            if (!piped && paste.Length > 0)
            {
                try { stream = new StreamReader(paste); }
                catch (Exception e) { Error("Unable to open file: " + paste + "\n\n"+errLine+"\n\n"+e.ToString(), 1); }
            }

            // number of lines and columns
            int N = 0;
            int C = 0;

            // input data, as numbers and text. only one will be used
            double[,] cells = null;
            string[,] sCells = null; 


            // if pasted text
            if (paste.Length > 0)
            {
                String line;
                List<string[]> strings = new List<string[]>();

                // trim all lines and split between words
                while ((line = stream.ReadLine()) != null)
                {
                    line = line.Replace(",", " ");
                    line = line.Replace("\t", " ");
                    line = line.Trim();
                    line = System.Text.RegularExpressions.Regex.Replace(line, @"\s+", " ");
                    if (asText)
                    {
                        //if (line.Length > 0)
                            strings.Add(line.Split(' '));
                    }
                    else if (line.Length > 0 && !line.StartsWith("#"))
                        strings.Add(line.Split(' '));
                }

                // set number of rows and columns
                N = strings.Count();
                C = 0;
                foreach (string[] str in strings) if (str.Length > C) C = str.Length;
                
                // convert input data to a format the Excel-interop understands: var[,]
                if (asText) sCells = new string[N, C];
                else cells = new double[N, C];

                // parse all cells and add to array
                for (int i = 0; i < N; i++)
                {
                    for (int j = 0; j < strings[i].Length; j++)
                    {
                        try
                        {
                            if (asText) sCells[i, j] = strings[i][j];
                            else cells[i, j] = double.Parse(strings[i][j]);
                        }
                        catch (Exception e) { Error("Unable to parse number in paste file, line " + (i + 1) + ", column " + (j + 1) + ":\n" + strings[i][j] + "\n\n" + errLine + "\n\n" + e.ToString(), 1); }
                    }
                }
            }


            Excel.Range startCell;
            Excel.Range endCell;

            try
            {
                // open excel app
                oXL = new Excel.Application();

                if (!warning) oXL.DisplayAlerts = false;

                try
                {
                    // try to open the selected excel file
                    // we turn of errors, since excel prompts for macro-enabled files and other things
                    oXL.DisplayAlerts = false;
                    // we need the absolute file path, since excel defaults to the user home dir, not the current working dir :S
                    oWB = oXL.Workbooks.Open(Path.GetFullPath(infile));
                    // turn back on warnings if wanted
                    if (warning) oXL.DisplayAlerts = true;
                }
                catch (Exception e) { throw new System.Exception("Unable to open file: " + Path.GetFullPath(infile) + "\n\n" + errLine + "\n\n" + e.ToString()); }

                // set the active sheet
                if (active.Length > 0)
                {
                    try
                    {
                        oSheet = (Excel._Worksheet)oWB.Sheets[active];
                    }
                    catch (Exception e) { throw new System.Exception("Unable to select worksheet: " + active + "\n\n" + errLine + "\n\n" + e.ToString()); }
                }
                else
                    oSheet = (Excel._Worksheet)oWB.Worksheets[1];


                // insert data
                if (paste.Length > 0)
                {
                    // if only columns are specified, find the amount of rows used
                    if (cellA[0] == 0 && cellB[0] == 0)
                    {
                        string cell = GetExcelColumnName(cellA[1]) + ":" + GetExcelColumnName(cellB[1]);
                        Excel.Range r = (Excel.Range)oSheet.UsedRange.Columns[cell, Type.Missing];
                        cellA[0] = 1;
                        cellB[0] = r.Rows.Count;
                    }

                    // select and paste values
                    try
                    {
                        startCell = (Excel.Range)oSheet.Cells[cellA[0], cellA[1]];
                        endCell = (Excel.Range)oSheet.Cells[cellB[0], cellB[1]];
                        oSheet.get_Range(startCell, endCell).Value = null;

                        endCell = (Excel.Range)oSheet.Cells[cellA[0] + N - 1, cellA[1] + C - 1];

                        if (asText) oSheet.get_Range(startCell, endCell).Value2 = sCells;
                        else oSheet.get_Range(startCell, endCell).Value2 = cells;
                    }
                    catch (Exception e)
                    {
                        string inputCell = GetExcelColumnName(cellA[1]) + cellA[0] + ":" + GetExcelColumnName(cellB[1]) + cellB[0];
                        throw new System.Exception("Unable to select input cells:\n\n         " + inputCell + "\n\n" + errLine + "\n\n" + e.ToString());
                    }
                }

                // run macro
                for (int i = 0; i < macro.Count; i++)
                {
                    try
                    {
                        oXL.Run(macro[i]);
                    }
                    catch (Exception e) { throw new System.Exception("Unable to run macro: " + macro[i] + "\n\n" + errLine + "\n\n" + e.ToString()); }
                }

                // force workbook refresh
                oXL.Calculate();

                // go to result sheet
                if (sheet.Length > 0)
                {
                    try
                    {
                        oSheet = (Excel._Worksheet)oWB.Sheets[sheet];
                    }
                    catch (Exception e) { throw new System.Exception("Unable to select output sheet:" + sheet + "\n\n" + errLine + "\n\n" + e.ToString()); }
                }

                // save charts 

                foreach (Excel.Worksheet cSheet in oWB.Worksheets)
                {
                    // loop trough all charts

                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)cSheet.ChartObjects(Type.Missing);
                    for (int i = 1; i <= xlCharts.Count; i++)
                    {

                        Excel.ChartObject oChart = (Excel.ChartObject)xlCharts.Item(i);
                        Excel.Chart chart = oChart.Chart;

                        string chartName = "";
                        if (charts.Contains(cSheet.Name + "." + oChart.Name)) chartName = cSheet.Name + "." + oChart.Name;
                        else if (charts.Contains(oChart.Name)) chartName = oChart.Name;

                        // if chart is specified for output, save it
                        if (chartName.Length > 0)
                        {
                            int id = charts.FindIndex(s => s == chartName);

                            charts.RemoveAt(id);

                            try
                            {
                                // we need full path name again.. excel defaults to user home dir...
                                string saveas = Path.GetFullPath(infile);
                                saveas = Path.GetDirectoryName(saveas) + "\\" + Path.GetFileNameWithoutExtension(saveas);
                                saveas = saveas + "_" + chartName + ".png";
                                chart.Export(saveas, "PNG");
                            }
                            catch (Exception e) { throw new System.Exception("Unable to save chart '" + chartName + "':\n\n" + errLine + "\n\n" + e.ToString()); }
                        }
                    }
                }

                // if any charts was not found; throw an error. 
                if (charts.Count > 0)
                {
                    string list = "";
                    foreach (string s in charts) list += s + ", ";
                    throw new Exception("Unable to find chart(s): " + list);
                }

                // if only columns are specified, find amount of rows to use
                if (outCellA[0] == 0 && outCellB[0] == 0)
                {

                    if (outCellB[1] == 0) outCellB[1] = oSheet.UsedRange.Columns.Count;

                    string cell = GetExcelColumnName(outCellA[1]) + ":" + GetExcelColumnName(outCellB[1]);
                    Excel.Range r = (Excel.Range)oSheet.UsedRange.Columns[cell, Type.Missing];
                    outCellA[0] = 1;
                    outCellB[0] = r.Rows.Count;
                }


                // select the output cell range
                try
                {
                    startCell = (Excel.Range)oSheet.Cells[outCellA[0], outCellA[1]];
                    endCell = (Excel.Range)oSheet.Cells[outCellB[0], outCellB[1]];
                }
                catch (Exception e)
                {
                    string outcell = GetExcelColumnName(outCellA[1]) + outCellA[0] + ":" + GetExcelColumnName(outCellB[1]) + outCellB[0];
                    throw new System.Exception("Unable to select output cells:\n            " + outcell + "\n\n" + errLine + "\n\n" + e.ToString());
                }


                // get output from selected cells
                object[,] arr = null;
                try
                {
                    Excel.Range r = (Excel.Range)oSheet.get_Range(startCell, endCell);
                    // if only 1 cell is selected, excel will return an object instead of object array!
                    if (r.Cells.Count == 1)
                    {
                        arr = new object[2, 2];
                        arr[1, 1] = r.Cells.Value2;
                    }
                    else arr = r.Cells.Value2 as object[,];
                }
                catch (Exception e)
                {
                    string outcell = GetExcelColumnName(outCellA[1]) + outCellA[0] + ":" + GetExcelColumnName(outCellB[1]) + outCellB[0];
                    throw new System.Exception("Invalid OutCells given. Unable to retrieve data:\n            " + outcell + "\n\n" + errLine + "\n\n" + e.ToString());
                }

                List<string> results = new List<string>();
                int last = 0;
                N = outCellB[0] - outCellA[0] + 1;
                C = outCellB[1] - outCellA[1] + 1;

                // loop trough output rows
                for (int i = 1; i <= N; i++)
                {
                    // loop trough output columns
                    string s = "";
                    for (int j = 1; j <= C; j++)
                    {
                        // check if cell contains an error 
                        if (arr[i, j] is Int32)
                        {
                            if (outErr)
                            {
                                int eCode = (int)arr[i, j];
                                string e = "";

                                if (eCode == -2146826281) e = "#DIV/0!";
                                else if (eCode == -2146826246) e = "#N/A";
                                else if (eCode == -2146826259) e = "#NAME?";
                                else if (eCode == -2146826288) e = "#NULL!";
                                else if (eCode == -2146826252) e = "#NUM!";
                                else if (eCode == -2146826265) e = "#REF!";
                                else if (eCode == -2146826273) e = "#VALUE!";
                                // no more error codes exists (?) as of 2013.. But to be sure / support future ones:
                                else e = "#ERR" + eCode.ToString();

                                s = s + e + " " + cellSpacer;
                            }
                            else s = s + " " + cellSpacer;
                        }
                        else if (arr[i, j] != null) s = s + arr[i, j].ToString() + cellSpacer;
                        else s = s + " " + cellSpacer;
                    }
                    // remove cellspacer from last column
                    if (C > 0) results.Add(s.Remove(s.Length - 1).TrimEnd());
                    // record last row column with content
                    if (s.Replace(cellSpacer, " ").TrimEnd().Length > 0) last = results.Count();
                }

                // write output to console
                for (int i = 0; i < last; i++)
                    Console.WriteLine(results[i]);

                // save file
                if (save)
                {
                    // if macros are enabled, excel would prompt about saving
                    oXL.DisplayAlerts = false;
                    oWB.Save();
                }

            }

                // catch any exception
            catch (Exception theException)
            {
                Error(errLine + "\n  Error: " + theException.Message, 1);
            }

            finally
            {

                // clean up and exit
                CleanUp();
            }


            #if DEBUG
                Console.ReadKey();
            #endif
        }


        // Release com the objects
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch 
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }


        // cleanup function
        public static void CleanUp()
        {
            try
            {
                oWB.Close(false);
                oXL.Quit();
                releaseObject(oSheet);
                releaseObject(oXL);
                releaseObject(oWB);
            }
            catch { }
        }

        // Print and exit functions
        public static void Error(string error, int ierr) { 
                Console.Error.WriteLine("ERROR " + error); 
                #if DEBUG
                   Console.ReadKey();
                #endif
                CleanUp();
                Environment.Exit(ierr); 
        }
        public static void Print(string str) { Console.WriteLine(str); }


        // convert excel cell reference to int [row,col]
        public static int[] ExcelCellRef(string str) {
                // remove $'locks, since they don't matter
                string cell = str.Replace("$", "");
                // convert RxCy style to x,y
                cell = System.Text.RegularExpressions.Regex.Replace(cell, @"R(\d+)C(\d+)", "$1,$2");
                // convert Xy style to X,y
                cell = System.Text.RegularExpressions.Regex.Replace(cell, @"([A-Za-z]+)(\d+)", "$1,$2");

                // replace letters with numbers
                Match match = Regex.Match(cell, @"^([A-Za-z]+)$", RegexOptions.IgnoreCase);
                if (match.Success) cell = cell+",0";

                try
                {
                    string[] s = cell.Split(',');
                    return new int[] { int.Parse(s[1]), ExcelColumnNameToNumber(s[0]) };
                }
                catch { 
                    throw new Exception("Unable to parse Cell: " + str + " (" + cell+ ")");
                }
        }

        // convert column letter to number
        public static int ExcelColumnNameToNumber(string columnName)
        {
            int sum;
            bool isNum = int.TryParse(columnName,out sum);
            if (isNum) return sum;
            columnName = columnName.ToUpperInvariant();
            sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum;
        }
        // convert column number to letter
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }



    static class ExcelMath
    {
        public static string Calc(string str)
        {
            string value = "";

            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application();
                xlApp.UseSystemSeparators = true;

                /*
                    string culture = System.Threading.Thread.CurrentThread.CurrentCulture.ToString();//"en-GB";
                    CultureInfo ci = new CultureInfo(culture);
                    Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("nb-NO");

                    xlApp.UseSystemSeparators = false;
                    xlApp.DecimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
                    xlApp.ThousandsSeparator = ci.NumberFormat.NumberGroupSeparator;

                    System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator = ";";
                */

                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                ws = (Excel.Worksheet)wb.Worksheets[1];

                //ws.get_Range("A1", "A1").Value2 = str;
                ws.get_Range("A1", "A1").FormulaLocal = str; // enter formula using the local culture

                xlApp.Calculate();
                value = ws.get_Range("A1", "A1").Value2.ToString();

                int eCode = 0;
                if (int.TryParse(value, out eCode))
                {
                    if (eCode == -2146826281) value = "#DIV/0!";
                    else if (eCode == -2146826246) value = "#N/A";
                    else if (eCode == -2146826259) value = "#NAME?";
                    else if (eCode == -2146826288) value = "#NULL!";
                    else if (eCode == -2146826252) value = "#NUM!";
                    else if (eCode == -2146826265) value = "#REF!";
                    else if (eCode == -2146826273) value = "#VALUE!";
                }

                wb.Close(false);
                xlApp.Quit();
            }

            catch (Exception ex)
            {
                value = "Error: " + ex.Message;
            }

            finally
            {

                try
                {
                    wb.Close(false);
                    xlApp.Quit();
                    releaseObject(ws);
                    releaseObject(wb);
                    releaseObject(xlApp);
                }
                catch
                {
                    releaseObject(ws);
                    releaseObject(wb);
                    releaseObject(xlApp);
                }

            }

            return value;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }



}
