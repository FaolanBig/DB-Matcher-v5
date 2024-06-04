/*
    The DB-Matcher is an easy-to-use console application written in C# and based on the .NET framework. 
    The DB-Matcher can merge two databases in Excel format (*.xlsx, *.xls). 
    It follows the following algorithms in order of importance: Levenshtein distance, Hamming distance, Jaccard index. 
    The DB-Matcher takes you by the hand at all times and guides you through the data matching process    

    Copyright (C) 2024  Carl Öttinger (Carl Oettinger)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as published
    by the Free Software Foundation, either version 3 of the License, or
    any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.

    You should have received a copy of the GNU Affero General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

    You can contact me in the following ways:
        EMail: oettinger.carl@web.de, big-programming@web.de
 */


using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;


namespace DB_Matching_main1
{
    internal class VarHold
    {
        public string getSimilarityMethodValue;
        /*public string GetSimilarityMethodValue
        {
            get { return getSimilarityMethodValue; }
            set { GetSimilarityMethodValue = value; }
        }*/
        //public BigInteger arrayAccess = BigInteger.Parse("0");
        public long arrayAccess = 0;
        public long cyclesLog = 0;
        public string checksumConvertedOriginal = "NULL";
        public string checksumConvertedNew = "NULL";
        public int runHold1 = 0;
        public int runHold2 = 0;
    }
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.Title = "DB-MATCHER";
            VarHold varHold = new VarHold();
            Stopwatch stopwatch = new Stopwatch();
        Start:
            printFittedSizeAsterixSurroundedText("DB-MATCHER");
            Console.WriteLine("### quit: ^C ###");
            Console.WriteLine();
        SetVerbose:
            Console.Write("Set Output to verbose (may be slower) y/n: ");
            bool verbose;
            /*if (Console.ReadKey(true).Key == ConsoleKey.Y)
            {
                verbose = true;
                Console.WriteLine("y");
            }
            else if (Console.ReadKey(true).Key == ConsoleKey.N)
            {
                verbose = false;
                Console.WriteLine("n");
            }
            else
            {
                Console.WriteLine();
                goto SetVerbose;
            }*/

            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    verbose = true;
                    Console.WriteLine("y"); break;
                case ConsoleKey.N:
                    verbose = false;
                    Console.WriteLine("n"); break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("DATA ERROR");
                    goto SetVerbose;
                    break;
            }

        SetWriteResults:
            bool writeResults;
            Console.WriteLine();
            Console.Write("Write Output (may be slower) y/n: ");
            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    writeResults = true;
                    Console.WriteLine("y");
                    break;
                case ConsoleKey.N:
                    writeResults = false;
                    Console.WriteLine("n");
                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("ERROR DATA");
                    goto SetWriteResults;
                    break;
            }

        ToggleConsoleBeep:
            bool toggleConsoleBeep;
            Console.WriteLine();
            Console.Write("Console Beep (may be annoying) y/n: ");
            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    toggleConsoleBeep = true;
                    Console.WriteLine("y");
                    Console.Beep();
                    break;
                case ConsoleKey.N:
                    toggleConsoleBeep = false;
                    Console.WriteLine("n");
                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("ERROR DATA");
                    goto ToggleConsoleBeep;
            }

        getSimilarityMethod:
            Console.WriteLine();
            Console.WriteLine("Similarity-Algorithm:\r\n...Levenshtein-Distance - 1\r\n...Hamming-Distance - 2\r\n...String-Contain - 3");
            Console.WriteLine();
            Console.Write("Press Key: ");

            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.D1:
                    varHold.getSimilarityMethodValue = "Levenshtein-Distance";
                    Console.WriteLine("1");
                    break;
                case ConsoleKey.D2:
                    varHold.getSimilarityMethodValue = "Hamming-Distance";
                    Console.WriteLine("2");
                    break;
                case ConsoleKey.D3:
                    varHold.getSimilarityMethodValue = "String-Contain";
                    Console.WriteLine("3");
                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("DATA ERROR");
                    goto getSimilarityMethod;
                    break;
            }
            Console.WriteLine();
        SetPath:
            Console.Write("Eingabe des Dateipfades: ");
            string path = Console.ReadLine(); //Pfad der Excel-Datei durch Konsoleneinabe
            //string path = $@"{pathHold}";
            Console.WriteLine();
            if (string.IsNullOrEmpty(path))
            {
                printFittedSizeAsterixSurroundedText("DATA ERROR");
                goto SetPath;
            }
            /*if (path.Contains('"'))
            {
                if (verbose) { Console.WriteLine("*** correcting path characters"); }
                path.Replace('"', ' ');
            }*/

            if (!File.Exists(path))
            {
                printFittedSizeAsterixSurroundedText("ERROR FILE NOT EXIST");
                path = null;
                goto SetPath;
            }

            /*FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite);
            workbook = new XSSFWorkbook(fileStream);*/

            printFittedSizeAsterixSurroundedText("COPY--FILE");

            int pathIndexOfDot = path.IndexOf('.');
            string toPath = path.Substring(0, pathIndexOfDot) + "_db-matched.xlsx";

        CheckIfToPathIsEmpty:
            varHold.runHold1 = 0;
            if (File.Exists(toPath))
            {
                varHold.runHold1++;
                //printFittedSizeAsterixSurroundedText("ERROR PATH NOT EMPTY");

                Random random = new Random();
                pathIndexOfDot = path.IndexOf('.');
                int randomLowerLimit = 10000;
                int randomUpperLimit = 100000;
                toPath = path.Substring(0, pathIndexOfDot) + $"_db-matched-{random.Next(randomLowerLimit, randomUpperLimit)}.xlsx";
                if (!(varHold.runHold1 > (randomUpperLimit - randomLowerLimit))) { goto CheckIfToPathIsEmpty; }
                else
                {
                    printFittedSizeAsterixSurroundedText("ERROR COPY TIMEOUT");
                    goto SetPath;
                }
            }

            try
            {
                File.Copy(path, toPath, false);
            }
            catch (Exception ex)
            {
                printFittedSizeAsterixSurroundedText("ERROR COPY");
                Console.WriteLine($"ERROR Message: {ex.Message}");
                goto CheckIfToPathIsEmpty;
            }
            Console.WriteLine("New Path: " + toPath);
            Console.WriteLine();

            printFittedSizeAsterixSurroundedText("STARTING PROCESS");

            IWorkbook workbook;

            using (var fs = new FileStream(toPath, FileMode.Open, FileAccess.Read)) //Lese-/Schreibzugriff
            {
                workbook = new XSSFWorkbook(fs);
            }

            if (toggleConsoleBeep) { Console.Beep(); }

        SheetInput:
            /*Console.Write("Arbeitsblatt (Nummer 1-32): ");
            int sheetInput = Convert.ToInt32(Console.ReadLine()) - 1;
            if (sheetInput < 0 || sheetInput > 31)
            {
                Console.WriteLine("Invalid Input");
                goto SheetInput;
            }
            ISheet sheet = workbook.GetSheetAt(sheetInput); //ExcelMappe Arbeitsblatt 0 (erstes)*/

            //Eingabe der Zellbereiche
            //printFittedSizeAsterixRow();
            Console.WriteLine();
        GetCellInput:
            Console.WriteLine("### Eingabe Bereich 1 ###");
            Console.WriteLine("### Eingabe in Zahlen ###");
        SheetInput1:
            Console.WriteLine();
            Console.Write("Arbeitsblatt (Nummer 1-32): ");
            int sheetInput1 = Convert.ToInt32(Console.ReadLine()) - 1;
            if (sheetInput1 < 0 || sheetInput1 > 31)
            {
                Console.WriteLine();
                printFittedSizeAsterixSurroundedText("DATA ERROR");
                goto GetCellInput;
            }
            Console.Write("Startzelle Bereich 1 Spaltenwert: ");
            int primaryFirstCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;

            Console.Write("Startzelle Bereich 1 Zeilenwert: ");
            int primaryFirstCellRow = Convert.ToInt32(Console.ReadLine()) - 1;

            Console.Write("Schlusszelle Bereich 1 Spaltenwert: ");
            int primaryLastCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;

            Console.Write("Schlusstzelle Bereich 1 Zeilenwert: ");
            int primaryLastCellRow = Convert.ToInt32(Console.ReadLine()) - 1;

            if (primaryFirstCellColumn == -1 || primaryFirstCellRow == -1 | primaryLastCellColumn == -1 | primaryLastCellRow == -1)
            {

                Console.WriteLine();
                printFittedSizeAsterixSurroundedText("ERROR DATA");
                Console.WriteLine();
                Console.WriteLine();
                goto GetCellInput;
            }

        GetCellInputSecondary:

            Console.WriteLine();
            Console.WriteLine("### Eingabe Bereich 2 ###");
            Console.WriteLine("### Eingabe in Zahlen ###");
        SheetInput2:
            Console.WriteLine();
            Console.Write("Arbeitsblatt (Nummer 1-32): ");
            int sheetInput2 = Convert.ToInt32(Console.ReadLine()) - 1;
            if (sheetInput2 < 0 || sheetInput2 > 31)
            {
                Console.WriteLine();
                printFittedSizeAsterixSurroundedText("DATA ERROR");
                goto GetCellInput;
            }


            Console.Write("Startzelle Bereich 2 Spaltenwert: ");
            int secondaryFirstCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;

            Console.Write("Startzelle Bereich 2 Zeilenwert: ");
            int secondaryFirstCellRow = Convert.ToInt32(Console.ReadLine()) - 1;

            Console.Write("Schlusszelle Bereich 2 Spaltenwert: ");
            int secondaryLastCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;

            Console.Write("Schlusstzelle Bereich 2 Zeilenwert: ");
            int secondaryLastCellRow = Convert.ToInt32(Console.ReadLine()) - 1;

            if (secondaryFirstCellColumn <= -1 || secondaryFirstCellRow <= -1 | secondaryLastCellColumn <= -1 | secondaryLastCellRow <= -1)
            {

                Console.WriteLine();
                printFittedSizeAsterixSurroundedText("ERROR DATA");
                Console.WriteLine();
                Console.WriteLine();
                goto GetCellInput;
            }

            if ((primaryFirstCellColumn != primaryLastCellColumn))
            {
                Console.WriteLine();

                string outputHold = "ERROR DATA";
                printFittedSizeAsterixSurroundedText(outputHold);
                /*int strLenght = outputHold.Length;
                
                for (int cnt = 0; cnt <= (Console.WindowWidth / 2 - strLenght / 2) - 1; cnt++)
                {
                    Console.Write("*");
                }
                Console.Write(outputHold);
                for (int cnt = 0; cnt <= (Console.WindowWidth / 2 - strLenght / 2) - 1; cnt++)
                {
                    Console.Write("*");
                }
                printFittedSizeAsterixRowCompact();
                Console.WriteLine();*/
                goto GetCellInput;
            }
            int resultSheet = 0;
            int resultColumn = 0;
            if (writeResults)
            {
                Console.WriteLine();
                Console.WriteLine("### Eingabe Bereich Ergebnis ###");
                Console.WriteLine("### Eingabe in Zahlen ###");
                Console.WriteLine();
                Console.Write("Arbeitsblatt (Nummer 1-32): ");
                resultSheet = Convert.ToInt32(Console.ReadLine()) - 1;
                Console.Write("Ergebnis Spaltenwert: ");
                resultColumn = Convert.ToInt32(Console.ReadLine()) - 1;

                if (resultColumn < 0 || resultSheet < 0)
                {
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("ERROR DATA");
                    Console.WriteLine();
                    Console.WriteLine();
                    goto GetCellInput;
                }
            }

            //Fortlaufendes Lesen der Zellen Zellbereich 1
            GetSimilarityValue getSimilarityValueOBJ = new GetSimilarityValue();
            Console.WriteLine();
            printFittedSizeAsterixSurroundedText($"STARTING COMPUTATION WITH {varHold.getSimilarityMethodValue} ALGORITHM");
            int matchedCells = 0;
            int matchedCellsIdentical = 0;
            if (verbose) { Console.WriteLine("starting stopwatch"); }
            stopwatch.Start();

            //Generating Checksums
            if (verbose) { printFittedSizeAsterixSurroundedText("COMPUTING CHECKSUM"); }
            using (FileStream fileStream = File.OpenRead(toPath))
            {
                SHA256Managed sha = new SHA256Managed();
                byte[] checksum = sha.ComputeHash(fileStream);
                varHold.checksumConvertedOriginal = BitConverter.ToString(checksum).Replace("-", String.Empty);
            }

            Stopwatch stopwatchIntern = new Stopwatch();
            stopwatchIntern.Start();
            FileStream ffstream = new FileStream(toPath, FileMode.Create, FileAccess.ReadWrite);
            for (int cnt = primaryFirstCellRow; cnt <= primaryLastCellRow; cnt++)
            {
                //stopwatchIntern.Start();
                ISheet sheet = workbook.GetSheetAt(sheetInput1);

                string activeMatchValue = "NOT FOUND";
                int activeMatchRow = 0;
                int activeMatchColumn = 0;
                double activeMatchPercentage = 100;

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

                    string activeSecondaryCellValue = activeSecondaryCell.ToString();

                    //Vergleichen
                    double compareMatchPercentage = getSimilarityValueOBJ.getLevenshteinDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    //double compareMatchPercentage = getSimilarityValueOBJ.getHammingDistance(activePrimaryCellValue, activeSecondaryCellValue);
                    if (activeSecondaryCellValue == activePrimaryCellValue)
                    {
                        matchedCells++;
                        matchedCellsIdentical++;
                        compareIdentical = true;
                        compareMatch = true;
                        compareMatchPercentage = 0;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchValue = activeSecondaryCellValue;
                        if (verbose == true) { Console.WriteLine(); printFittedSizeAsterixSurroundedText("MATCH--FOUND"); }
                    }
                    else if (compareMatchPercentage < activeMatchPercentage)
                    {
                        matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        compareMatch = true;
                        //compareMatchPercentage = getSimilarityValue(activePrimaryCellValue, activeSecondaryCellValue);
                        if (verbose == true) { Console.WriteLine(); printFittedSizeAsterixSurroundedText("MATCH--FOUND"); }
                    }
                    else
                    {
                        compareMatch = false;
                    }


                    if (verbose == true)
                    {
                        if (compareMatch) printFittedSizeAsterixRowCompact();
                        /*Console.Write("Active Computing: ");
                        Console.Write("Matched Cells: " + matchedCells + " | ");
                        Console.Write("Matched Cells Identical: " + matchedCellsIdentical + " | ");
                        Console.Write("Active Primary Row: " + (cnt + 1) + " | ");
                        Console.Write("Active Secondary Row: " + (sCnt + 1) + " | ");
                        Console.Write("ActivePrimaryCellValue: " + activePrimaryCellValue + " | ");
                        Console.Write("ActiveSecondaryCellValue: " + activeSecondaryCellValue + " | ");
                        Console.Write("CompareIdentical: " + compareIdentical + " | ");
                        Console.Write("CompareMatch: " + compareMatch + " | ");
                        Console.Write("CompareMatchPercentage: " + compareMatchPercentage);
                        Console.WriteLine();*/

                        string outputHold = "";
                        outputHold += ("Active Computing: ");
                        outputHold += ("Matched Cells: " + matchedCells + " | ");
                        outputHold += ("Matched Cells Identical: " + matchedCellsIdentical + " | ");
                        outputHold += ("Active Primary Row: " + (cnt + 1) + " | ");
                        outputHold += ("Active Secondary Row: " + (sCnt + 1) + " | ");
                        outputHold += ("ActivePrimaryCellValue: " + activePrimaryCellValue + " | ");
                        outputHold += ("ActiveSecondaryCellValue: " + activeSecondaryCellValue + " | ");
                        outputHold += ("CompareIdentical: " + compareIdentical + " | ");
                        outputHold += ("CompareMatch: " + compareMatch + " | ");
                        outputHold += ("CompareMatchPercentage: " + compareMatchPercentage);
                        Console.WriteLine(outputHold);

                        if (compareMatch) printFittedSizeAsterixRowCompact();
                        /*
                        progress = cnt / (primaryLastCellRow - primaryFirstCellRow);
                        string progressHold = "";
                        string progressOutputHold = "";
                        Console.SetCursorPosition(0, Console.WindowHeight - 1);
                        Console.Write(new string(' ', Console.WindowWidth));
                        Console.SetCursorPosition(0, Console.WindowHeight - 1);
                        for (int i = 0; i < (Console.WindowWidth - (Console.WindowWidth - progress)); i++)
                        {
                            progressOutputHold += "#";
                        }
                        Console.Write(progressOutputHold);*/
                    }
                    /*progress = cnt / primaryLastCellRow;
                    double progressHold = 0.0;
                    string progressOutputHold = "";
                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                    Console.Write(new string(' ', Console.WindowWidth));
                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                    /*for (int i = 0; i < (Console.WindowWidth - (Console.WindowWidth - progress)); i++)
                    {
                        progressOutputHold += "#";
                    }
                    for (int i = 0; i < Convert.ToInt32(Console.WindowWidth * progress); i++)
                    {
                        progressOutputHold += "#";
                    }
                    Console.Write(progressOutputHold);
                    Thread.Sleep(100);*/
                    varHold.cyclesLog++;
                }
                //if (verbose) Console.WriteLine();
                if (!verbose)
                {
                    //stopwatchIntern.Stop();
                    TimeSpan timeSpanIntern = stopwatchIntern.Elapsed;
                    string timeSpanStringIntern = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);
                    //double timeSpanInternDouble = stopwatchIntern.Elapsed.TotalMilliseconds / ((cnt - primaryFirstCellRow) + 1);
                    //string timeSpanStringInternViaDouble = String.Format("{0:00}:{1:00}:{2:00}", Total, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                    double timeSpanMillisecondsDoubleHold = stopwatchIntern.ElapsedMilliseconds;
                    double timeSpanMillisecondsPerIterationHold = timeSpanMillisecondsDoubleHold / ((cnt - primaryFirstCellRow) + 1);
                    int totalIterations = primaryLastCellRow - primaryFirstCellRow;
                    double remainingTime = timeSpanMillisecondsPerIterationHold * (totalIterations - (cnt - primaryFirstCellRow));

                    TimeSpan timeSpanRemainingTimeFormatted = TimeSpan.FromMilliseconds(remainingTime);
                    string timeSpanRemainingTimeStringFormatted = string.Format("{0:D2}h:{1:D2}m:{2:D2}s", timeSpanRemainingTimeFormatted.Hours, timeSpanRemainingTimeFormatted.Minutes, timeSpanRemainingTimeFormatted.Seconds);

                    /*progress = (cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow);
                    float progressHold = cnt / primaryLastCellRow;
                    string progressOutputHold = "";*/
                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                    Console.Write(new string(' ', Console.WindowWidth));
                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                    /*for (int i = 0; i < (Console.WindowWidth - (Console.WindowWidth - progress)); i++)
                    {
                        progressOutputHold += "#";
                    }*/
                    /*for (int i = 0; i < (Console.WindowWidth * progressHold); i++)
                    {
                        //progressOutputHold += "#";
                        Console.Write("#");
                    }
                    //Console.Write(progressOutputHold);*/
                    /*int tprogress = (cnt * (primaryLastCellRow - primaryFirstCellRow) / primaryLastCellRow);
                    Console.Write($"[{new string('#', tprogress)}{new string(' ', (primaryLastCellRow - primaryFirstCellRow) - tprogress)}] {cnt}%");*/
                    /*int tprogress = (cnt * (Console.WindowWidth - 2) / primaryLastCellRow);
                    Console.Write($"[{new string('#', tprogress)}{new string(' ', (Console.WindowWidth - 2) - tprogress)}]");*/
                    int progressPercentage = Convert.ToInt32(Convert.ToDouble(cnt - primaryFirstCellRow) / (primaryLastCellRow - primaryFirstCellRow) * 100);

                    string remainingTimeString = String.Format("{0:00}:{1:00}:{2:00}", timeSpanIntern.Hours, timeSpanIntern.Minutes, timeSpanIntern.Seconds);

                    double tprogress = (cnt * (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) / primaryLastCellRow);
                    Console.Write($"[{new string('#', Convert.ToInt32(tprogress))}{new string(' ', (Console.WindowWidth - (9 + timeSpanRemainingTimeStringFormatted.Length)) - Convert.ToInt32(tprogress))}]{progressPercentage}% | {timeSpanRemainingTimeStringFormatted}");

                    //Thread.Sleep(100);
                }
                //Gefundene Werte in definierten Bereich schreiben
                if (writeResults)
                {
                    //using (FileStream ffstream = new FileStream(toPath, FileMode.Create, FileAccess.ReadWrite))
                    //{
                    string cellValueTransferHold;
                    for (int i = secondaryFirstCellColumn; i <= secondaryLastCellColumn; i++)
                    {
                        //read
                        string outputHold = null;
                        if (verbose) { outputHold = "*** prepare for copy ***"; }
                        if (verbose) { Console.WriteLine(outputHold); }

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
                        if (verbose) { outputHold = "*** copy ***"; }
                        if (verbose) { Console.WriteLine(outputHold); }
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
                    sheet = workbook.GetSheetAt(resultSheet);
                    IRow rrow = sheet.GetRow(cnt);
                    ICell ccell = rrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + 1);
                    cellValueTransferHold = $"LD-Value: {activeMatchPercentage}";
                    ccell.SetCellValue(cellValueTransferHold);

                    //workbook.Write(ffstream);
                    //ffstream.Close();
                    //}
                }
            }
            //Console.WriteLine("schreiben");
            Console.WriteLine();
            Console.WriteLine();
            printFittedSizeAsterixSurroundedText("SAVING");

            workbook.Write(ffstream);
            ffstream.Close();

            if (verbose) { printFittedSizeAsterixSurroundedText("COMPUTING CHECKSUM"); }
            using (FileStream fffstream = File.OpenRead(toPath))
            {
                SHA256Managed sha = new SHA256Managed();
                byte[] checksum = sha.ComputeHash(fffstream);
                varHold.checksumConvertedNew = BitConverter.ToString(checksum).Replace("-", String.Empty);
            }

            if (toggleConsoleBeep) { Console.Beep(); }

            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            string timeSpanString = String.Format("{0:00}h:{1:00}m:{2:00}s:{3:00}ms", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds, timeSpan.Milliseconds);

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Matched Cells: " + matchedCells);
            Console.WriteLine("Matched Cells Identical: " + matchedCellsIdentical);
            Console.WriteLine();
            Console.WriteLine("Old Checksum (SHA256): " + varHold.checksumConvertedOriginal);
            Console.WriteLine("New Checksum (SHA256): " + varHold.checksumConvertedNew);
            Console.WriteLine();
            Console.WriteLine("Cycles: " + varHold.cyclesLog);
            Console.WriteLine("ArrayAccessLog: " + getSimilarityValueOBJ.getArrayAccess());
            Console.WriteLine();
            Console.WriteLine("Computing Duration: " + timeSpanString);

            //exit
            /*using (FileStream ffstream = new FileStream(path, FileMode.Open, FileAccess.Write))
            {
                workbook.Write(ffstream);
            }*/
            //fileStream.Dispose();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Press key: ENTER to restart in new instance / ESC to exit");
        GetExitValue:
            bool toggleRestartInNewInstance = false;
            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Enter:
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("RESTARTING IN NEW INSTANCE");
                    workbook.Close();
                    workbook = null;
                    toggleRestartInNewInstance = true;
                    goto exit;
                    break;
                case ConsoleKey.Escape:
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("SHUTDOWN");
                    goto exit;
                    break;
                default:
                    goto GetExitValue;
                    break;
            }
        exit:
            Console.WriteLine();
            for (int i = 0; i <= 100; i++)
            {
                Console.SetCursorPosition(0, Console.WindowHeight - 1);
                Console.Write(new string(' ', Console.WindowWidth));
                Console.SetCursorPosition(0, Console.WindowHeight - 1);

                double tprogress = (i * (Console.WindowWidth - 6) / 100);
                Console.Write($"[{new string('#', Convert.ToInt32(tprogress))}{new string(' ', (Console.WindowWidth - 6) - Convert.ToInt32(tprogress))}]{i}%");
                Thread.Sleep(1);
            }
            if (toggleRestartInNewInstance)
            {
                var currentFileName = Assembly.GetExecutingAssembly().Location;
                Process.Start(currentFileName);
            }
            Environment.Exit(0);
        }

        private static void printFittedSizeAsterixRow()
        {
            Console.WriteLine();
            for (int cnt = 0; cnt < Console.WindowWidth; cnt++)
            {
                Console.Write("*");
            }
            Console.WriteLine();
            Console.WriteLine();
        }

        private static void printFittedSizeAsterixRowCompact()
        {
            string outputHold = "";
            for (int cnt = 0; cnt < Console.WindowWidth; cnt++)
            {
                outputHold += "*";
            }
            Console.WriteLine(outputHold);
            //Console.WriteLine();
        }

        private static void printFittedSizeHashtagRow()
        {
            Console.WriteLine();
            for (int cnt = 0; cnt < Console.WindowWidth; cnt++)
            {
                Console.Write("#");
            }
            Console.WriteLine();
            Console.WriteLine();
        }

        private static void printFittedSizeHashtagRowCompact()
        {
            for (int cnt = 0; cnt < Console.WindowWidth; cnt++)
            {
                Console.Write("#");
            }
            Console.WriteLine();
        }

        private static void printFittedSizeAsterixSurroundedText(string text)
        {
            int strLenght = text.Length;
            string outputHold = "";

            printFittedSizeAsterixRowCompact();
            for (int cnt = 0; cnt <= (Console.WindowWidth / 2 - strLenght / 2) - 2; cnt++)
            {
                //Console.Write("*");
                outputHold += "*";
            }
            //Console.Write(" " + text + " ");
            outputHold += $" {text} ";
            for (int cnt = 0; cnt <= (Console.WindowWidth / 2 - strLenght / 2) - 2; cnt++)
            {
                //Console.Write("*");
                outputHold += "*";
            }
            Console.WriteLine(outputHold);
            printFittedSizeAsterixRowCompact();
            Console.WriteLine();
        }
    }
    public class GetSimilarityValue
    {
        private int[,] matrixHold = new int[2, 2];
        private VarHold varhold = new VarHold();
        public bool toggleArrayAccessLog = true;

        public double getLevenshteinDistance(string stringHold1, string stringHold2) //Levenshtein-Distance Algorithmus
        {
            int lenthHold1 = stringHold1.Length;
            int lenthHold2 = stringHold2.Length;

            matrixHold = new int[lenthHold1 + 1, lenthHold2 + 1];
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }

            for (int i = 0; i <= lenthHold1; i++)
            {
                matrixHold[i, 0] = i;
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            }
            for (int j = 0; j <= lenthHold2; j++)
            {
                //Console.WriteLine("LenthHold: " + lenthHold2);
                //Console.WriteLine("  j: " + j);
                matrixHold[0, j] = j;
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            }
            for (int i = 1; i <= lenthHold1; i++)
            {
                for (int j = 1; j <= lenthHold2; j++)
                {
                    /*int val = (stringHold2[j - 1] == stringHold1[i - 1]) ? 0 : 1;
                    matrixHold[i, j] = Math.Min(Math.Min(matrixHold[i - 1, j] + 1, matrixHold[i, j - 1] + 1), matrixHold[i - 1, j - 1] + val);*/
                    int cost = (j - 1 < stringHold2.Length && i - 1 < stringHold1.Length && stringHold2[j - 1] == stringHold1[i - 1]) ? 0 : 1;
                    matrixHold[i, j] = Math.Min(Math.Min(matrixHold[i - 1, j] + 1, matrixHold[i, j - 1] + 1), matrixHold[i - 1, j - 1] + cost);
                    if (toggleArrayAccessLog) { varhold.arrayAccess += 3; }
                }
            }

            //double returnHold = matrixHold[lenthHold1, lenthHold2];
            //matrixHold = null;
            //double returnHold = 1.0 - (double)matrixHold[lenthHold1, lenthHold2] / Math.Max(lenthHold1, lenthHold2);
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            return matrixHold[lenthHold1, lenthHold2];
            //return returnHold;
        }
        public int getHammingDistance(string stringHold1, string stringHold2)
        {
            prepareGetHammingDistance(stringHold1, stringHold2);
            int hammingDistance = 0;
            for (int i = 0; i < stringHold1.Length; i++)
            {
                if (stringHold1[i] != stringHold2[i]) { hammingDistance++; }
            }

            return (hammingDistance);
        }
        private bool prepareGetHammingDistance(string stringHold1, string stringHold2)
        {
            if (stringHold1.Length == stringHold2.Length)
            {
                return (true);
            }
            else
            {
                adaptStringLength(stringHold1, stringHold2);
            }

            return (false);
        }
        private bool adaptStringLength(string stringHold1, string stringHold2)
        {
            if (stringHold1.Length == stringHold2.Length) { return (false); }
            else if (stringHold1.Length < stringHold2.Length)
            {
                while (stringHold1.Length < stringHold2.Length)
                {
                    stringHold1 += null; //Prüfen
                }
            }
            else
            {
                while (stringHold1.Length > stringHold2.Length)
                {
                    stringHold2 += null;
                }
            }

            return (true);
        }
        public short getStringContainingValue(string stringHold1, string stringHold2) //0: Equal; 1: 2 in 1; 2: 2 in 1; 3: no
        {
            if (stringHold1.Equals(stringHold2)) { return 0; }
            else if (stringHold1.Contains(stringHold2)) { return 1; }
            else if (stringHold2.Contains(stringHold1)) { return 2; }
            else { return 3; }
        }
        public long getArrayAccess()
        {
            return (varhold.arrayAccess);
        }
    }
}
