using DB_Matching_main1;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
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
            int errorCount = 0;
            PrintIn.blue("running preparation");

            if (File.Exists(VarHold.helperFilePath))
            {
                PrintIn.yellow("previously used helperFile detected");
                RecoveryHandler.WaitForKeystrokeENTER("hit ENTER to delete the helperFile and continue");
                File.Delete(VarHold.helperFilePath);
            }

            ISheet tsheet = workbook.GetSheetAt(sheet);
            using (StreamWriter writer = new StreamWriter(VarHold.helperFilePath))
            {
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
            if (errorCount == 0) { PrintIn.green("preparation finished successful"); }
            else { PrintIn.yellow($"preparation finished with {errorCount} minor errors"); PrintIn.yellow("see log for more information"); }
        }
        internal static void writeFileToContent(ref IWorkbook workbook, int sheet, int primaryFirstCellColumn, int primaryFirstCellRow, int primaryLastCellColumn, int primaryLastCellRow)
        {
            int errorCount = 0;

            PrintIn.blue("cleaning");

            if (!File.Exists(VarHold.helperFilePath))
            {
                PrintIn.red("no helperFile found");
                PrintIn.red("no cleaning possible");
                RecoveryHandler.WaitForKeystrokeENTER("hit ENTER to continue without cleaning");
            }

            ISheet tsheet = workbook.GetSheetAt(sheet);
            using (StreamReader reader = new(VarHold.helperFilePath))
            {
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
            File.Delete(VarHold.helperFilePath);

            if (errorCount == 0) { PrintIn.green("cleaning finished successful"); }
            else { PrintIn.yellow($"preparation finished with {errorCount} minor errors"); PrintIn.yellow("see log for more information"); }
        }

    }
}
