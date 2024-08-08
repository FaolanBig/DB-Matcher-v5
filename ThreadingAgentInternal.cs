using DB_Matching_main1;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal class ThreadingAgentInternal
    {
        internal static void Matcher(IWorkbook workbook, short threadCount, ICellStyle[] styles, Dictionary<string, string> dictionary, int sheetInput1, int sheetInput2, int resultSheet, int resultColumn, int primaryFirstCellRow, int primaryLastCellRow, int primaryFirstCellColumn, int primaryLastCellColumn, int secondaryFirstCellRow, int secondaryLastCellRow, int secondaryFirstCellColumn, int secondaryLastCellColumn)
        {
            GetSimilarityValue getSimilarityValueOBJ = new GetSimilarityValue();
            Stopwatch stopwatchIntern = new Stopwatch();
            stopwatchIntern.Start();

            for (int cnt = primaryFirstCellRow; cnt <= primaryLastCellRow; cnt++)
            {
                //stopwatchIntern.Start();
                ISheet sheet = workbook.GetSheetAt(sheetInput1);

                string activeMatchValue = "NOT FOUND";
                int activeMatchRow = 0;
                int activeMatchColumn = 0;
                double activeMatchPercentage = 100;
                int activeMatchHammingDistance = 100;
                double activeMatchJaccardIndex = 1.0;

                string activeSecondaryCellValue = "NOT FOUND";
                string activePrimaryCellValueOld = null;
                string activeSecondaryCellValueOld = null;

                IRow activePrimaryCellRow = sheet.GetRow(cnt);
                if (activePrimaryCellRow == null)
                {
                    activePrimaryCellRow = sheet.CreateRow(cnt);
                }
                ICell activePrimaryCell = activePrimaryCellRow.GetCell(primaryFirstCellColumn);
                if (activePrimaryCell == null)
                {
                    activePrimaryCell = activePrimaryCellRow.CreateCell(primaryFirstCellColumn);
                }

                string activePrimaryCellValue = activePrimaryCell.ToString();

                if (VarHold.useDataFile)
                {
                    foreach (var entry in dictionary)
                    {
                        if (activePrimaryCellValue.Contains(entry.Key))
                        {
                            activePrimaryCellValueOld = activePrimaryCellValue;
                            activePrimaryCellValue = activePrimaryCellValue.Replace(entry.Key, entry.Value);
                            break;
                        }
                    }
                }

                for (int sCnt = secondaryFirstCellRow; sCnt <= secondaryLastCellRow; sCnt++)
                {
                    sheet = workbook.GetSheetAt(sheetInput2);

                    bool compareIdentical = false;
                    bool compareMatch = false;


                    IRow activeSecondaryCellRow = sheet.GetRow(sCnt);
                    if (activeSecondaryCellRow == null)
                    {
                        activeSecondaryCellRow = sheet.CreateRow(sCnt);
                    }
                    ICell activeSecondaryCell = activeSecondaryCellRow.GetCell(secondaryFirstCellColumn);
                    if (activeSecondaryCell == null)
                    {
                        activeSecondaryCell = activeSecondaryCellRow.CreateCell(secondaryFirstCellColumn);
                    }

                    activeSecondaryCellValue = activeSecondaryCell.ToString();

                    if (VarHold.useDataFile)
                    {
                        foreach (var entry in dictionary)
                        {
                            if (activeSecondaryCellValue.Contains(entry.Key))
                            {
                                activeSecondaryCellValueOld = activeSecondaryCellValue;
                                activeSecondaryCellValue = activeSecondaryCellValue.Replace(entry.Key, entry.Value);
                                break;
                            }
                        }
                    }

                    //Vergleichen
                    double compareMatchPercentage = getSimilarityValueOBJ.getLevenshteinDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    int compareMatchHammingDistance = getSimilarityValueOBJ.getHammingDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    double compareMatchJaccardIndex = getSimilarityValueOBJ.getJaccardIndex(activePrimaryCellValue, activeSecondaryCellValue);
                    //double compareMatchPercentage = getSimilarityValueOBJ.getHammingDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    if (activeSecondaryCellValue == activePrimaryCellValue)
                    {
                        VarHold.matchedCells++;
                        VarHold.matchedCellsIdentical++;
                        compareIdentical = true;
                        compareMatch = true;
                        compareMatchPercentage = 0;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchValue = activeSecondaryCellValue;
                    }
                    else if (compareMatchPercentage < activeMatchPercentage)
                    {
                        VarHold.matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        compareMatch = true;
                        //compareMatchPercentage = getSimilarityValue(activePrimaryCellValue, activeSecondaryCellValue);
                    }
                    /*else if (compareMatchPercentage == activeMatchPercentage && compareMatchJaccardIndex < activeMatchJaccardIndex)
                    {
                        matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        //activeMatchPercentage = compareMatchPercentage;
                        activeMatchJaccardIndex = compareMatchJaccardIndex;
                    }*/
                    else if (compareMatchPercentage == activeMatchPercentage && compareMatchHammingDistance < activeMatchHammingDistance)
                    {
                        VarHold.matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchHammingDistance = compareMatchHammingDistance;
                    }
                    else if (compareMatchPercentage == activeMatchPercentage && compareMatchHammingDistance == activeMatchHammingDistance && compareMatchJaccardIndex < activeMatchJaccardIndex)
                    {
                        VarHold.matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchJaccardIndex = compareMatchJaccardIndex;
                        activeMatchJaccardIndex = compareMatchJaccardIndex;

                    }
                    else
                    {
                        compareMatch = false;
                    }


                    VarHold.cyclesLog++;
                }
                if (VarHold.writeResults)
                {
                    //using (FileStream ffstream = new FileStream(toPath, FileMode.Create, FileAccess.ReadWrite))
                    //{
                    string cellValueTransferHold;
                    for (int i = secondaryFirstCellColumn; i <= secondaryLastCellColumn; i++)
                    {
                        //read
                        string outputHold = null;

                        sheet = workbook.GetSheetAt(sheetInput2);
                        IRow fromRow = sheet.GetRow(activeMatchRow);
                        if (fromRow == null)
                        {
                            fromRow = sheet.CreateRow(cnt);
                        }
                        ICell fromCell = fromRow.GetCell(i);
                        if (fromCell == null)
                        {
                            fromCell = activePrimaryCellRow.CreateCell(primaryFirstCellColumn);
                        }
                        cellValueTransferHold = fromCell.ToString();

                        //write
                        //Console.WriteLine(1);
                        sheet = workbook.GetSheetAt(resultSheet);
                        //Console.WriteLine(2);
                        IRow toRow = sheet.GetRow(cnt);
                        //Console.WriteLine(3);
                        ICell toCell = toRow.CreateCell(resultColumn + (i - secondaryFirstCellColumn));
                        //Console.WriteLine(4);
                        //toCell.SetCellValue(cellValueTransferHold);
                        toCell.SetCellValue(cellValueTransferHold);
                        //Console.WriteLine(5);

                        /*using (FileStream ffstream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                        {
                            workbook.Write(ffstream);
                        }*/
                        //workbook.Write(fileStream);
                    }
                    //Console.WriteLine(6);
                    //workbook.Write(ffstream);
                    //Console.WriteLine(7);
                    //ffstream.Close();
                    //}
                    //using (FileStream ffstream = new FileStream(toPath, FileMode.Create, FileAccess.Write))
                    //{
                    //Console.WriteLine("Anhang schreiben");
                    int columnHold = 0;

                    sheet = workbook.GetSheetAt(resultSheet);
                    IRow rrrow = sheet.GetRow(cnt);
                    ICell cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"LD-Value: {activeMatchPercentage}";
                    cccell.SetCellValue(cellValueTransferHold);
                    if (SettingsAgent.GetSettingIsTrue("colorGradient"))
                    {
                        if (Convert.ToInt32(activeMatchPercentage) < 10) { cccell.CellStyle = styles[Convert.ToInt32(activeMatchPercentage)]; }
                        else { cccell.CellStyle = styles[10]; }
                    }

                    sheet = workbook.GetSheetAt(resultSheet);
                    rrrow = sheet.GetRow(cnt);
                    cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"HD-Value: {activeMatchHammingDistance}";
                    cccell.SetCellValue(cellValueTransferHold);

                    sheet = workbook.GetSheetAt(resultSheet);
                    rrrow = sheet.GetRow(cnt);
                    cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"JD-Value: {activeMatchJaccardIndex}";
                    cccell.SetCellValue(cellValueTransferHold);

                    if (VarHold.useDataFile && (activePrimaryCellValueOld != null || activeSecondaryCellValueOld != null))
                    {
                        sheet = workbook.GetSheetAt(resultSheet);
                        IRow rrow = sheet.GetRow(cnt);
                        ICell ccell = rrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                        cellValueTransferHold = $"PrimaryValueTransform (old --> new): {activePrimaryCellValueOld} --> {activePrimaryCellValue}";
                        ccell.SetCellValue(cellValueTransferHold);

                        sheet = workbook.GetSheetAt(resultSheet);
                        rrow = sheet.GetRow(cnt);
                        ccell = rrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                        cellValueTransferHold = $"SecondaryValueTransform (old --> new): {activeSecondaryCellValueOld} --> {activeSecondaryCellValue}";
                        ccell.SetCellValue(cellValueTransferHold);
                    }

                    //workbook.Write(ffstream);
                    //ffstream.Close();
                    //}
                }

                TimeSpan timeSpanIntern = stopwatchIntern.Elapsed;
                string timeSpanStringIntern = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                double timeSpanMillisecondsDoubleHold = stopwatchIntern.ElapsedMilliseconds;
                double timeSpanMillisecondsPerIterationHold = timeSpanMillisecondsDoubleHold / ((cnt - primaryFirstCellRow) + 1);
                int totalIterations = primaryLastCellRow - primaryFirstCellRow;
                double remainingTime = timeSpanMillisecondsPerIterationHold * (totalIterations - (cnt - primaryFirstCellRow));

                TimeSpan timeSpanRemainingTimeFormatted = TimeSpan.FromMilliseconds(remainingTime);
                string timeSpanRemainingTimeStringFormatted = string.Format("{0:D2}h:{1:D2}m:{2:D2}s", timeSpanRemainingTimeFormatted.Hours, timeSpanRemainingTimeFormatted.Minutes, timeSpanRemainingTimeFormatted.Seconds);

                double tprogress = (cnt * (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) / primaryLastCellRow);
                int progressPercentage = Convert.ToInt32(Convert.ToDouble(cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow) * 100);

                /*if (threadCount == 1) { VarHold.thread1_progress = Convert.ToInt32(((cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow)) * 100); VarHold.thread1_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 2) { VarHold.thread2_progress = Convert.ToInt32(((cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow)) * 100); VarHold.thread2_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 3) { VarHold.thread3_progress = Convert.ToInt32(((cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow)) * 100); VarHold.thread3_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 4) { VarHold.thread4_progress = Convert.ToInt32(((cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow)) * 100); VarHold.thread4_remainingTime = timeSpanRemainingTimeStringFormatted; }
                */
                if (threadCount == 1) { VarHold.thread1_progress = tprogress; }
                else if (threadCount == 2) { VarHold.thread2_progress = tprogress; }
                else if (threadCount == 3) { VarHold.thread3_progress = tprogress; }
                else if (threadCount == 4) { VarHold.thread4_progress = tprogress; }

                if (threadCount == 1) { VarHold.thread1_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 2) { VarHold.thread2_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 3) { VarHold.thread3_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 4) { VarHold.thread4_remainingTime = timeSpanRemainingTimeStringFormatted; }

            if (threadCount == 1) { VarHold.thread1_progressPercentage = progressPercentage; }
                else if (threadCount == 2) { VarHold.thread2_progressPercentage = progressPercentage; }
                else if (threadCount == 3) { VarHold.thread3_progressPercentage = progressPercentage; }
                else if (threadCount == 4) { VarHold.thread4_progressPercentage = progressPercentage; }

            }
        }
        internal static void Matcher_objTransfer(dataTransferHoldObj obj)
        {
            GetSimilarityValue getSimilarityValueOBJ = new GetSimilarityValue();
            Stopwatch stopwatchIntern = new Stopwatch();
            stopwatchIntern.Start();

            for (int cnt = obj.fromPrimary; cnt <= obj.toPrimary; cnt++)
            {
                //stopwatchIntern.Start();
                ISheet sheet = workbook.GetSheetAt(sheetInput1);

                string activeMatchValue = "NOT FOUND";
                int activeMatchRow = 0;
                int activeMatchColumn = 0;
                double activeMatchPercentage = 100;
                int activeMatchHammingDistance = 100;
                double activeMatchJaccardIndex = 1.0;

                string activeSecondaryCellValue = "NOT FOUND";
                string activePrimaryCellValueOld = null;
                string activeSecondaryCellValueOld = null;

                IRow activePrimaryCellRow = sheet.GetRow(cnt);
                if (activePrimaryCellRow == null)
                {
                    activePrimaryCellRow = sheet.CreateRow(cnt);
                }
                ICell activePrimaryCell = activePrimaryCellRow.GetCell(obj.primaryColumn);
                if (activePrimaryCell == null)
                {
                    activePrimaryCell = activePrimaryCellRow.CreateCell(obj.primaryColumn);
                }

                string activePrimaryCellValue = activePrimaryCell.ToString();

                if (VarHold.useDataFile)
                {
                    foreach (var entry in dictionary)
                    {
                        if (activePrimaryCellValue.Contains(entry.Key))
                        {
                            activePrimaryCellValueOld = activePrimaryCellValue;
                            activePrimaryCellValue = activePrimaryCellValue.Replace(entry.Key, entry.Value);
                            break;
                        }
                    }
                }

                for (int sCnt = secondaryFirstCellRow; sCnt <= secondaryLastCellRow; sCnt++)
                {
                    sheet = workbook.GetSheetAt(sheetInput2);

                    bool compareIdentical = false;
                    bool compareMatch = false;


                    IRow activeSecondaryCellRow = sheet.GetRow(sCnt);
                    if (activeSecondaryCellRow == null)
                    {
                        activeSecondaryCellRow = sheet.CreateRow(sCnt);
                    }
                    ICell activeSecondaryCell = activeSecondaryCellRow.GetCell(secondaryFirstCellColumn);
                    if (activeSecondaryCell == null)
                    {
                        activeSecondaryCell = activeSecondaryCellRow.CreateCell(secondaryFirstCellColumn);
                    }

                    activeSecondaryCellValue = activeSecondaryCell.ToString();

                    if (VarHold.useDataFile)
                    {
                        foreach (var entry in dictionary)
                        {
                            if (activeSecondaryCellValue.Contains(entry.Key))
                            {
                                activeSecondaryCellValueOld = activeSecondaryCellValue;
                                activeSecondaryCellValue = activeSecondaryCellValue.Replace(entry.Key, entry.Value);
                                break;
                            }
                        }
                    }

                    //Vergleichen
                    double compareMatchPercentage = getSimilarityValueOBJ.getLevenshteinDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    int compareMatchHammingDistance = getSimilarityValueOBJ.getHammingDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    double compareMatchJaccardIndex = getSimilarityValueOBJ.getJaccardIndex(activePrimaryCellValue, activeSecondaryCellValue);
                    //double compareMatchPercentage = getSimilarityValueOBJ.getHammingDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    if (activeSecondaryCellValue == activePrimaryCellValue)
                    {
                        VarHold.matchedCells++;
                        VarHold.matchedCellsIdentical++;
                        compareIdentical = true;
                        compareMatch = true;
                        compareMatchPercentage = 0;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchValue = activeSecondaryCellValue;
                    }
                    else if (compareMatchPercentage < activeMatchPercentage)
                    {
                        VarHold.matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        compareMatch = true;
                        //compareMatchPercentage = getSimilarityValue(activePrimaryCellValue, activeSecondaryCellValue);
                    }
                    /*else if (compareMatchPercentage == activeMatchPercentage && compareMatchJaccardIndex < activeMatchJaccardIndex)
                    {
                        matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        //activeMatchPercentage = compareMatchPercentage;
                        activeMatchJaccardIndex = compareMatchJaccardIndex;
                    }*/
                    else if (compareMatchPercentage == activeMatchPercentage && compareMatchHammingDistance < activeMatchHammingDistance)
                    {
                        VarHold.matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchHammingDistance = compareMatchHammingDistance;
                    }
                    else if (compareMatchPercentage == activeMatchPercentage && compareMatchHammingDistance == activeMatchHammingDistance && compareMatchJaccardIndex < activeMatchJaccardIndex)
                    {
                        VarHold.matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchJaccardIndex = compareMatchJaccardIndex;
                        activeMatchJaccardIndex = compareMatchJaccardIndex;

                    }
                    else
                    {
                        compareMatch = false;
                    }


                    VarHold.cyclesLog++;
                }
                if (VarHold.writeResults)
                {
                    //using (FileStream ffstream = new FileStream(toPath, FileMode.Create, FileAccess.ReadWrite))
                    //{
                    string cellValueTransferHold;
                    for (int i = secondaryFirstCellColumn; i <= secondaryLastCellColumn; i++)
                    {
                        //read
                        string outputHold = null;

                        sheet = workbook.GetSheetAt(sheetInput2);
                        IRow fromRow = sheet.GetRow(activeMatchRow);
                        if (fromRow == null)
                        {
                            fromRow = sheet.CreateRow(cnt);
                        }
                        ICell fromCell = fromRow.GetCell(i);
                        if (fromCell == null)
                        {
                            fromCell = activePrimaryCellRow.CreateCell(obj.primaryColumn);
                        }
                        cellValueTransferHold = fromCell.ToString();

                        //write
                        //Console.WriteLine(1);
                        sheet = workbook.GetSheetAt(resultSheet);
                        //Console.WriteLine(2);
                        IRow toRow = sheet.GetRow(cnt);
                        //Console.WriteLine(3);
                        ICell toCell = toRow.CreateCell(resultColumn + (i - secondaryFirstCellColumn));
                        //Console.WriteLine(4);
                        //toCell.SetCellValue(cellValueTransferHold);
                        toCell.SetCellValue(cellValueTransferHold);
                        //Console.WriteLine(5);

                        /*using (FileStream ffstream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                        {
                            workbook.Write(ffstream);
                        }*/
                        //workbook.Write(fileStream);
                    }
                    //Console.WriteLine(6);
                    //workbook.Write(ffstream);
                    //Console.WriteLine(7);
                    //ffstream.Close();
                    //}
                    //using (FileStream ffstream = new FileStream(toPath, FileMode.Create, FileAccess.Write))
                    //{
                    //Console.WriteLine("Anhang schreiben");
                    int columnHold = 0;

                    sheet = workbook.GetSheetAt(resultSheet);
                    IRow rrrow = sheet.GetRow(cnt);
                    ICell cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"LD-Value: {activeMatchPercentage}";
                    cccell.SetCellValue(cellValueTransferHold);
                    if (SettingsAgent.GetSettingIsTrue("colorGradient"))
                    {
                        if (Convert.ToInt32(activeMatchPercentage) < 10) { cccell.CellStyle = styles[Convert.ToInt32(activeMatchPercentage)]; }
                        else { cccell.CellStyle = styles[10]; }
                    }

                    sheet = workbook.GetSheetAt(resultSheet);
                    rrrow = sheet.GetRow(cnt);
                    cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"HD-Value: {activeMatchHammingDistance}";
                    cccell.SetCellValue(cellValueTransferHold);

                    sheet = workbook.GetSheetAt(resultSheet);
                    rrrow = sheet.GetRow(cnt);
                    cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"JD-Value: {activeMatchJaccardIndex}";
                    cccell.SetCellValue(cellValueTransferHold);

                    if (VarHold.useDataFile && (activePrimaryCellValueOld != null || activeSecondaryCellValueOld != null))
                    {
                        sheet = workbook.GetSheetAt(resultSheet);
                        IRow rrow = sheet.GetRow(cnt);
                        ICell ccell = rrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                        cellValueTransferHold = $"PrimaryValueTransform (old --> new): {activePrimaryCellValueOld} --> {activePrimaryCellValue}";
                        ccell.SetCellValue(cellValueTransferHold);

                        sheet = workbook.GetSheetAt(resultSheet);
                        rrow = sheet.GetRow(cnt);
                        ccell = rrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                        cellValueTransferHold = $"SecondaryValueTransform (old --> new): {activeSecondaryCellValueOld} --> {activeSecondaryCellValue}";
                        ccell.SetCellValue(cellValueTransferHold);
                    }

                    //workbook.Write(ffstream);
                    //ffstream.Close();
                    //}
                }

                TimeSpan timeSpanIntern = stopwatchIntern.Elapsed;
                string timeSpanStringIntern = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                double timeSpanMillisecondsDoubleHold = stopwatchIntern.ElapsedMilliseconds;
                double timeSpanMillisecondsPerIterationHold = timeSpanMillisecondsDoubleHold / ((cnt - obj.fromPrimary) + 1);
                int totalIterations = obj.toPrimary - obj.fromPrimary;
                double remainingTime = timeSpanMillisecondsPerIterationHold * (totalIterations - (cnt - obj.fromPrimary));

                TimeSpan timeSpanRemainingTimeFormatted = TimeSpan.FromMilliseconds(remainingTime);
                string timeSpanRemainingTimeStringFormatted = string.Format("{0:D2}h:{1:D2}m:{2:D2}s", timeSpanRemainingTimeFormatted.Hours, timeSpanRemainingTimeFormatted.Minutes, timeSpanRemainingTimeFormatted.Seconds);

                double tprogress = (cnt * (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) / obj.toPrimary);
                int progressPercentage = Convert.ToInt32(Convert.ToDouble(cnt - obj.fromPrimary) / (obj.toPrimary - obj.fromPrimary) * 100);

                /*if (threadCount == 1) { VarHold.thread1_progress = Convert.ToInt32(((cnt - obj.fromPrimary) / (obj.toPrimary - obj.fromPrimary)) * 100); VarHold.thread1_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 2) { VarHold.thread2_progress = Convert.ToInt32(((cnt - obj.fromPrimary) / (obj.toPrimary - obj.fromPrimary)) * 100); VarHold.thread2_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 3) { VarHold.thread3_progress = Convert.ToInt32(((cnt - obj.fromPrimary) / (obj.toPrimary - obj.fromPrimary)) * 100); VarHold.thread3_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 4) { VarHold.thread4_progress = Convert.ToInt32(((cnt - obj.fromPrimary) / (obj.toPrimary - obj.fromPrimary)) * 100); VarHold.thread4_remainingTime = timeSpanRemainingTimeStringFormatted; }
                */
                if (threadCount == 1) { VarHold.thread1_progress = tprogress; }
                else if (threadCount == 2) { VarHold.thread2_progress = tprogress; }
                else if (threadCount == 3) { VarHold.thread3_progress = tprogress; }
                else if (threadCount == 4) { VarHold.thread4_progress = tprogress; }

                if (threadCount == 1) { VarHold.thread1_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 2) { VarHold.thread2_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 3) { VarHold.thread3_remainingTime = timeSpanRemainingTimeStringFormatted; }
                else if (threadCount == 4) { VarHold.thread4_remainingTime = timeSpanRemainingTimeStringFormatted; }

                if (threadCount == 1) { VarHold.thread1_progressPercentage = progressPercentage; }
                else if (threadCount == 2) { VarHold.thread2_progressPercentage = progressPercentage; }
                else if (threadCount == 3) { VarHold.thread3_progressPercentage = progressPercentage; }
                else if (threadCount == 4) { VarHold.thread4_progressPercentage = progressPercentage; }

            }
        }

    }
}
