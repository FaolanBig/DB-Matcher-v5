using DB_Matching_main1;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal static class Helper
    {
        internal static void writeContentToFile(IWorkbook workbook, int sheet, int primaryFirstCellColumn, int primaryFirstCellRow, int primaryLastCellColumn, int primaryLastCellRow)
        {
            ToLog.Inf("writing content to file @Helper.cs");
            int errorCount = 0;
            PrintIn.blue("running preparation");

            if (File.Exists(VarHold.helperFilePath))
            {
                ToLog.Err("previously used helperFile detected");
                PrintIn.yellow("previously used helperFile detected");
                RecoveryHandler.WaitForKeystrokeENTER("hit ENTER to delete the helperFile and continue");
                ToLog.Inf($"deleting file: {VarHold.helperFilePath}");
                File.Delete(VarHold.helperFilePath);
                ToLog.Inf("file deleted: success");
            }

            ISheet tsheet = workbook.GetSheetAt(sheet);
            using (StreamWriter writer = new StreamWriter(VarHold.helperFilePath))
            {
                ToLog.Inf("starting stopwatch(intern) @Helper.cs");
                Stopwatch stopwatchIntern = new Stopwatch();

                for (int row = primaryFirstCellRow; row <= primaryLastCellRow; row++)
                {
                    IRow currentRow = tsheet.GetRow(row);
                    if (currentRow != null)
                    {
                        ICell cell = currentRow.GetCell(primaryFirstCellColumn);
                        if (cell != null && cell.CellType == CellType.String)
                        {
                            writer.WriteLine(cell.StringCellValue);
                        }
                        else
                        {
                            ToLog.Err($"error during preparation - row empty: {row}");
                            errorCount++;
                            PrintIn.red("an error occurred during preparation");
                            PrintIn.red("see log for more information");
                            RecoveryHandler.WaitForKeystrokeENTER();
                            writer.WriteLine(" ");
                            //Program.shutdownOrRestart();
                        }
                    }
                    else
                    {
                        ToLog.Err($"error during preparation - row empty: {row}");
                        PrintIn.red("an error occurred during preparation");
                        PrintIn.red("see log for more information");
                        writer.WriteLine(" ");
                    }

                    TimeSpan timeSpanIntern = stopwatchIntern.Elapsed;
                    string timeSpanStringIntern = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                    double timeSpanMillisecondsDoubleHold = stopwatchIntern.ElapsedMilliseconds;
                    double timeSpanMillisecondsPerIterationHold = timeSpanMillisecondsDoubleHold / ((row - primaryFirstCellRow) + 1);
                    int totalIterations = primaryLastCellRow - primaryFirstCellRow;
                    double remainingTime = timeSpanMillisecondsPerIterationHold * (totalIterations - (row - primaryFirstCellRow));

                    TimeSpan timeSpanRemainingTimeFormatted = TimeSpan.FromMilliseconds(remainingTime);
                    string timeSpanRemainingTimeStringFormatted = string.Format("{0:D2}h:{1:D2}m:{2:D2}s", timeSpanRemainingTimeFormatted.Hours, timeSpanRemainingTimeFormatted.Minutes, timeSpanRemainingTimeFormatted.Seconds);

                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                    Console.Write(new string(' ', Console.WindowWidth));
                    Console.SetCursorPosition(0, Console.WindowHeight - 1);

                    int progressPercentage = Convert.ToInt32(Convert.ToDouble(row - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow) * 100);

                    string remainingTimeString = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                    double tprogress = (row * (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) / primaryLastCellRow);
                    Console.Write($"[{new string('#', Convert.ToInt32(tprogress))}{new string('.', (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) - Convert.ToInt32(tprogress))}]{progressPercentage}% | {timeSpanRemainingTimeStringFormatted}");

                }
            }
            if (errorCount == 0) { ToLog.Inf("operation has finished: writing content to file"); PrintIn.green("preparation finished successful"); }
            else { ToLog.Err($"operation has finished: writing content to file - {errorCount} minor errors occurred"); PrintIn.yellow($"preparation finished with {errorCount} minor errors"); PrintIn.yellow("see log for more information"); }
        }
        internal static void readContentFromFile(ref IWorkbook workbook, int sheet, int primaryFirstCellColumn, int primaryFirstCellRow, int primaryLastCellColumn, int primaryLastCellRow)
        {
            int errorCount = 0;

            PrintIn.blue("cleaning");

            if (!File.Exists(VarHold.helperFilePath))
            {
                ToLog.Err("no file found: helper file - abort");
                PrintIn.red("no helperFile found");
                PrintIn.red("no cleaning possible");
                RecoveryHandler.WaitForKeystrokeENTER("hit ENTER to continue without cleaning");
            }

            ISheet tsheet = workbook.GetSheetAt(sheet);
            using (StreamReader reader = new(VarHold.helperFilePath))
            {
                ToLog.Inf("starting stopwatch(intern)");
                Stopwatch stopwatchIntern = new Stopwatch();

                string line;
                int row = primaryFirstCellRow;
                while ((line = reader.ReadLine()) != null)
                {
                    IRow currentRow = tsheet.GetRow(row) ?? tsheet.CreateRow(row);
                    ICell newCell = currentRow.CreateCell(primaryFirstCellColumn); // Neue Spalte
                    newCell.SetCellValue(line);

                    try
                    {
                        TimeSpan timeSpanIntern = stopwatchIntern.Elapsed;
                        string timeSpanStringIntern = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                        double timeSpanMillisecondsDoubleHold = stopwatchIntern.ElapsedMilliseconds;
                        double timeSpanMillisecondsPerIterationHold = timeSpanMillisecondsDoubleHold / ((row - primaryFirstCellRow) + 1);
                        int totalIterations = primaryLastCellRow - primaryFirstCellRow;
                        double remainingTime = timeSpanMillisecondsPerIterationHold * (totalIterations - (row - primaryFirstCellRow));
                        if (double.IsInfinity(remainingTime) || double.IsNaN(remainingTime)) {remainingTime = 0;}

                        TimeSpan timeSpanRemainingTimeFormatted = TimeSpan.FromMilliseconds(remainingTime);
                        string timeSpanRemainingTimeStringFormatted = string.Format("{0:D2}h:{1:D2}m:{2:D2}s", timeSpanRemainingTimeFormatted.Hours, timeSpanRemainingTimeFormatted.Minutes, timeSpanRemainingTimeFormatted.Seconds);

                        Console.SetCursorPosition(0, Console.WindowHeight - 1);
                        Console.Write(new string(' ', Console.WindowWidth));
                        Console.SetCursorPosition(0, Console.WindowHeight - 1);

                        int progressPercentage = Convert.ToInt32(Convert.ToDouble(row) / Convert.ToDouble(totalIterations) * 100);

                        string remainingTimeString = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                        double tprogress = (row * (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) / primaryLastCellRow);
                        //double tprogress = Convert.ToInt32(row) / Convert.ToDouble(totalIterations);
                        Console.Write($"[{new string('#', 1 + Convert.ToInt32(tprogress))}{new string('.', (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) - 1 - Convert.ToInt32(tprogress))}]{progressPercentage}% | {timeSpanRemainingTimeStringFormatted}");
                    }
                    catch (Exception ex)
                    {
                        ToLog.Err($"an error occurred when calculating or printing progress (progress bar) - error: {ex.Message}");
                        PrintIn.red("an error occurred when calculating or printing progress (progress bar)");
                        PrintIn.red("see log for more information");
                        RecoveryHandler.WaitForKeystrokeENTER();
                    }

                    row++;
                }
            }
            ToLog.Inf("deleting file: helper file");
            File.Delete(VarHold.helperFilePath);
            ToLog.Inf("helper file deleted: success");

            if (errorCount == 0) { ToLog.Inf("operation has finished: reading content from file"); PrintIn.green("cleaning finished successful"); }
            else { ToLog.Err($"operation has finished: reading content from file - {errorCount} minor errors occurred"); PrintIn.yellow($"preparation finished with {errorCount} minor errors"); PrintIn.yellow("see log for more information"); }
        }
        internal static void printProgressBar(int line, double progress, int total, int barCount, string remainingTime, int progressPercentage)
        {
            int progressBarLenth = 50;
            Console.SetCursorPosition(0, Console.WindowHeight - (barCount - line));

            Console.Write($"[{new string('#', Convert.ToInt32(progress))}{new string('.', (Console.WindowWidth - (9 + remainingTime.Length)) - Convert.ToInt32(progress))}]{progressPercentage}% | {remainingTime} \n");
        }
        internal static void DrawProgressBar(int line, int progress, int total, int barCount, string remainingTime)
        {
            //Console.SetCursorPosition(0, Console.WindowHeight - (barCount - line + 1));
            //Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, Console.WindowHeight - (barCount - line + 1));
            string outputHold = $"thread{line}: [";
            Console.Write($"thread{line}: [");
            int pos = 50 * progress / total;
            for (int i = 0; i < 50; i++)
            {
                if (i < pos) {Console.Write("="); outputHold += "="; }
                else if (i == pos) {Console.Write(">"); outputHold += ">"; }
                else {Console.Write(" "); outputHold += " "; }
            }
            outputHold += $"] {progress}% | {remainingTime}";
            //outputHold += new string(' ', Console.WindowWidth - 1 - outputHold.Length) + "\n";
            Console.Write($"] {progress}% | {remainingTime}");
            Console.SetCursorPosition(outputHold.Length, Console.WindowHeight - (barCount - line + 1));
            Console.Write(new string(' ', 5));
            //Console.Write(outputHold);
        }
        internal static string returnProgressBar(int line, int progress, int total, int barCount, string remainingTime)
        {
            string returnHold = $"thread{line}: [";

            int pos = 50 * progress / total;
            for (int i = 0; i < 50; i++)
            {
                if (i < pos) returnHold += "=";
                else if (i == pos) returnHold += ">";
                else returnHold += " ";
            }
            returnHold += $"] {progress}% | {remainingTime}";
            return returnHold;
        }
        internal static void openURLinDefaultBrowser(string url = VarHold.repoURLReportIssue)
        {
            PrintIn.blue($"loading URL in default browser: {url}");

            try
            {
                Process.Start("cmd", $"/C start {url}");
            }
            catch (System.ComponentModel.Win32Exception noDefaultBrowser)
            {
                ToLog.Err($"error when launching URL in default browser: {noDefaultBrowser.Message}");
                PrintIn.red("error when launching URL in default browser");
                PrintIn.red("see log for more information");
            }
            catch (Exception ex)
            {
                ToLog.Err($"error when launching URL in default browser: {ex.Message}");
                PrintIn.red($"error when launching URL in default browser");
                PrintIn.red("see log for more information");
            }
        }
    }
}
