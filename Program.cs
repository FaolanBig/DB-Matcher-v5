﻿/*
   The DB-Matcher (including DB-Matcher-v5) is an easy-to-use console application written in C# and based on the .NET framework. 
   The DB-Matcher can merge two databases in Excel format (*.xlsx, *.xls). 
   It follows the following algorithms in order of importance: Levenshtein distance, Hamming distance, Jaccard index. 
   The DB-Matcher takes you by the hand at all times and guides you through the process of data matching. 

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
       EMail: oettinger.carl@web.de or big-programming@web.de
*/


using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;


namespace DB_Matching_main1
{
    internal class VarHold
    {
        public string currentMainFilePath;
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
        public string createJsonExitInput = "JSONEXIT";
        public string createDictionaryExitInput = "DIREXIT";
        public bool useDataFile = true;
        public bool toggleConsoleColor = true;
    }
    internal class Program
    {
        private static void Main(string[] args)
        {
            run();
        }

        private static void run(bool jsonCheck = true)
        {
            Console.Title = "DB-MATCHER-v5";
            VarHold varHold = new VarHold();
            Stopwatch stopwatch = new Stopwatch();
            Dictionary<string, string> dictionary = new Dictionary<string, string>();

        Start:
            printFittedSizeAsterixSurroundedText("DB-MATCHER");

            string currentHoldFilePath = AppDomain.CurrentDomain.BaseDirectory;
            string currentHoldFilePath2 = currentHoldFilePath;
            string currentSettingsFilePathHold = currentHoldFilePath;
            varHold.currentMainFilePath = Assembly.GetExecutingAssembly().Location;
            currentHoldFilePath += "data.txt";
            currentHoldFilePath2 += "pw.txt";
            currentSettingsFilePathHold += "settings.txt";

            //StartUp Interrupt
            if (Console.KeyAvailable)
            {
                if (Console.ReadKey(true).Key == ConsoleKey.Escape)
                {
                    Console.WriteLine("Interrupt during StartUp");
                    Console.WriteLine("Press ENTER to continue");

                    while (Console.ReadKey(true).Key != ConsoleKey.Enter) { }

                    Console.WriteLine();
                }
            }

        ContinueFromInterruptDuringStartUp:
            if (File.Exists(currentHoldFilePath2))
            {
                string pwHold = "";
                bool makeHash = false;

                try
                {
                    using (StreamReader sReader = new StreamReader(currentHoldFilePath2))
                    {
                        pwHold = sReader.ReadLine();
                        if (!string.IsNullOrEmpty(pwHold)) { makeHash = true; }
                    }
                }
                catch (Exception)
                {
                    File.Delete(currentHoldFilePath2);
                }

                if (makeHash)
                {
                    SHA256 sha = SHA256Managed.Create();
                    byte[] bytes = Encoding.UTF8.GetBytes(pwHold);
                    byte[] data = sha.ComputeHash(bytes);

                    StringBuilder result = new StringBuilder();

                    for (int i = 0; i < data.Length; i++)
                    {
                        result.Append(data[i].ToString("X2"));
                    }

                    string fromHash = result.ToString();
                    string toHash = "E0D40F7F4E95E42FBBDC773CC08C998CE6A13550DD3621B7F3DE3D512B879864";

                    if (fromHash == toHash)
                    {
                        string writeHold007 = @"                   XXXX
                  X    XX
                 X  ***  X                XXXXX
                X  *****  X            XXX     XX
             XXXX ******* XXX      XXXX          XX
           XX	X ******  XXXXXXXXX    El@         XX XXX
         XX	 X ****  X                           X** X
        X        XX    XX     X                      X***X
       X         //XXXX       X                      XXXX
      X         //   X                             XX
     X         //    X	        XXXXXXXXXXXXXXXXXX/
     X	   XXX//    X          X
     X	  X   X     X         X
     X    X    X    X        X
      X   X    X    X        X			  XX
      X    X   X    X        X		       XXX  XX
       X    XXX      X        X 	      X  X X  X
       X	     X         X	      XX X  XXXX
        X	      X         XXXXXXXX\     XX   XX  X
         XX	       XX             X     X    @X  XX
           XX		 XXXX	XXXXXX/     X     XXXX
             XXX	     XX***         X     X
                XXXXXXXXXXXXX *   *       X     X
                             *---* X     X     X
                            *-* *   XXX X     X
                            *- *       XXX   X
                           *- *X	  XXX
                           *- *X  X	     XXX
                          *- *X    X		XX
                          *- *XX    X		  X
                         *  *X* X    X		   X
                         *  *X * X    X 	    X
                        *  * X**  X   XXXX	    X
                        *  * X**  XX	 X	    X
                       *  ** X** X     XX	   X
                       *  **  X*  XXX	X	  X
                      *  **    XX   XXXX       XXX
                     *	* *	 XXXX	   X	 X
                    *	* *	     X	   X	 X
      =======*******   * *	     X	   X	  XXXXXXXX\
             *	       * *	/XXXXX	    XXXXXXXX\	   )
        =====**********  *     X		     )	\  )
          ====* 	*     X 	      \  \   )XXXXX
     =========**********       XXXXXXXXXXXXXXXXXXXXXX";
                        Console.WriteLine(writeHold007);
                        Console.WriteLine();
                    }
                }
            }

            if (jsonCheck)
            {
                if (!File.Exists(currentHoldFilePath))
                {
                    createDictionary(currentHoldFilePath);

                    //neu starten
                    Process process = new Process();
                    process.StartInfo.FileName = "cmd.exe";
                    process.StartInfo.Arguments = $"/C {varHold.currentMainFilePath}";

                    printFittedSizeAsterixSurroundedText("RESTARTING");

                    for (int i = 0; i <= 100; i++)
                    {
                        Console.SetCursorPosition(0, Console.WindowHeight - 1);
                        Console.Write(new string(' ', Console.WindowWidth));
                        Console.SetCursorPosition(0, Console.WindowHeight - 1);

                        double tprogress = (i * (Console.WindowWidth - 6) / 100);
                        Console.Write($"[{new string('#', Convert.ToInt32(tprogress))}{new string(' ', (Console.WindowWidth - 6) - Convert.ToInt32(tprogress))}]{i}%");
                        Thread.Sleep(1);
                    }

                    process.Start();
                    Environment.Exit(1);
                }

                Console.WriteLine("### quit: ^C ###");
                Console.WriteLine();

                using (StreamReader file = new StreamReader(currentHoldFilePath))
                {
                    string line;
                    while ((line = file.ReadLine()) != null)
                    {
                        string[] streamParts = line.Split(' ');
                        if (streamParts.Length == 2)
                        {
                            dictionary.Add(streamParts[0], streamParts[1]);
                        }
                    }
                }
                Console.WriteLine("Current Dictionary Content: ");

                foreach (var entry in dictionary)
                {
                    Console.WriteLine("- {0} ---> {1}", entry.Key, entry.Value);
                }

                Console.WriteLine();
            KeepDataFile:
                Console.Write("Keep DATA file (y/n): ");

                switch (Console.ReadKey(true).Key)
                {
                    case ConsoleKey.Y:
                        Console.WriteLine("y");
                        break;
                    case ConsoleKey.N:
                        Console.WriteLine("n");
                    ConfirmDelete:
                        Console.Write("Confirm delete DATA file (y/n): ");
                        switch (Console.ReadKey(true).Key)
                        {
                            case ConsoleKey.Y:
                                Console.WriteLine("y");
                                File.Delete(currentHoldFilePath);
                                Console.WriteLine();
                                setConsoleColorToGreen();
                                Console.WriteLine("*** deleted ***");
                                resetConsoleColor();
                                Console.WriteLine();
                                printFittedSizeAsterixSurroundedText("RESTARTING");

                                Process process = new Process();
                                process.StartInfo.FileName = "cmd.exe";
                                process.StartInfo.Arguments = $"/C {varHold.currentMainFilePath}";

                                for (int i = 0; i <= 100; i++)
                                {
                                    Console.SetCursorPosition(0, Console.WindowHeight - 1);
                                    Console.Write(new string(' ', Console.WindowWidth));
                                    Console.SetCursorPosition(0, Console.WindowHeight - 1);

                                    double tprogress = (i * (Console.WindowWidth - 6) / 100);
                                    Console.Write($"[{new string('#', Convert.ToInt32(tprogress))}{new string(' ', (Console.WindowWidth - 6) - Convert.ToInt32(tprogress))}]{i}%");
                                    Thread.Sleep(1);
                                }

                                process.Start();
                                Environment.Exit(1);

                                break;
                            case ConsoleKey.N:
                                Console.WriteLine("n");
                                Console.WriteLine();
                                setConsoleColorToRed();
                                Console.WriteLine("*** canceled ***");
                                resetConsoleColor();
                                break;
                            default:
                                printFittedSizeAsterixSurroundedText("ERROR DATA");
                                goto ConfirmDelete;
                                break;
                        }
                        break;
                    default:
                        Console.WriteLine();
                        printFittedSizeAsterixSurroundedText("ERROR DATA");
                        goto KeepDataFile;
                        break;
                }

                Console.WriteLine();

            UseDataFile:
                Console.Write("Use DATA file(y/n): ");
                switch (Console.ReadKey(true).Key)
                {
                    case ConsoleKey.Y:
                        Console.WriteLine("y");
                        varHold.useDataFile = true;
                        break;
                    case ConsoleKey.N:
                        Console.WriteLine("n");
                        varHold.useDataFile = false;
                        break;
                    default:
                        goto UseDataFile;
                        break;
                }
            }

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

        /*ToggleConsoleColor:
            Console.WriteLine();
            Console.Write("Console Color (y/n): ");
            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    varHold.toggleConsoleColor = true;
                    Console.WriteLine("y");
                    Console.Beep();
                    break;
                case ConsoleKey.N:
                    varHold.toggleConsoleColor = false;
                    Console.WriteLine("n");
                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("ERROR DATA");
                    goto ToggleConsoleBeep;
            }*/


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
            Console.WriteLine();

        getSimilarityMethod:
            varHold.getSimilarityMethodValue = "Levenshtein-Distance";
        /*Console.WriteLine();
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
        Console.WriteLine();*/
        SetPath:
            Console.Write("Eingabe des Dateipfades: ");
            string path = Console.ReadLine(); //Pfad der Excel-Datei durch Konsoleneinabe
            path = path.Replace("\"", "");
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
            Console.Write("Arbeitsblatt (Nummer 1-32, default 1): ");
            //string inputHold1 = Console.ReadLine();
            int sheetInput1 = getIntOrDefault(0);
            //if (!string.IsNullOrEmpty(inputHold1)) { sheetInput1 = Convert.ToInt32(inputHold1) - 1; }

            if (sheetInput1 < 0 || sheetInput1 > 31)
            {
                Console.WriteLine();
                printFittedSizeAsterixSurroundedText("DATA ERROR");
                goto GetCellInput;
            }

            Console.Write("Startzelle Bereich 1 Spaltenwert (default: 1): ");
            //int primaryFirstCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;
            //inputHold1 = Console.ReadLine();
            int primaryFirstCellColumn = getIntOrDefault(0);
            //if (!string.IsNullOrEmpty(inputHold1)) { primaryFirstCellColumn = Convert.ToInt32(inputHold1) - 1; }


            Console.Write("Startzelle Bereich 1 Zeilenwert (default: 1): ");
            //int primaryFirstCellRow = Convert.ToInt32(Console.ReadLine()) - 1;
            int primaryFirstCellRow = getIntOrDefault(0);

            Console.Write($"Schlusszelle Bereich 1 Spaltenwert (default: {primaryFirstCellColumn + 1}): ");
            //int primaryLastCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;
            int primaryLastCellColumn = getIntOrDefault(primaryFirstCellColumn);

            Console.Write($"Schlusszelle Bereich 1 Zeilenwert (default: {getLastRowWithValue(workbook, sheetInput1, primaryFirstCellColumn) + 1}): ");
            var holdConsoleInput1 = Console.ReadLine();
            int primaryLastCellRow = 0;
            if (string.IsNullOrEmpty(holdConsoleInput1)) { primaryLastCellRow = getLastRowWithValue(workbook, sheetInput1, primaryFirstCellColumn); }
            else { primaryLastCellRow = Convert.ToInt32(holdConsoleInput1) - 1; }

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
            Console.Write($"Arbeitsblatt (Nummer 1-32, default: {sheetInput1 + 1}): ");
            //int sheetInput2 = Convert.ToInt32(Console.ReadLine()) - 1;
            int sheetInput2 = getIntOrDefault(sheetInput1);
            if (sheetInput2 < 0 || sheetInput2 > 31)
            {
                Console.WriteLine();
                printFittedSizeAsterixSurroundedText("DATA ERROR");
                goto GetCellInput;
            }

            int defaultHold = 0;
            if (sheetInput1 != sheetInput2) { defaultHold = primaryFirstCellColumn; }
            else { defaultHold = primaryFirstCellColumn + 1; }

            Console.Write($"Startzelle Bereich 2 Spaltenwert (default: {defaultHold + 1}): ");
            //int secondaryFirstCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;
            int secondaryFirstCellColumn = getIntOrDefault(defaultHold);

            Console.Write($"Startzelle Bereich 2 Zeilenwert (default: {primaryFirstCellRow + 1}): ");
            //int secondaryFirstCellRow = Convert.ToInt32(Console.ReadLine()) - 1;
            int secondaryFirstCellRow = getIntOrDefault(primaryFirstCellRow);

            Console.Write($"Schlusszelle Bereich 2 Spaltenwert (default: {getLastColumnWithValue(workbook, sheetInput2, secondaryFirstCellRow) + 1}): ");
            //int secondaryLastCellColumn = Convert.ToInt32(Console.ReadLine()) - 1;
            int secondaryLastCellColumn = getIntOrDefault(getLastColumnWithValue(workbook, sheetInput2, secondaryFirstCellRow));

            Console.Write($"Schlusszelle Bereich 2 Zeilenwert (default: {getLastRowWithValue(workbook, sheetInput2, secondaryFirstCellColumn) + 1}): ");
            var holdConsoleInput2 = Console.ReadLine();
            int secondaryLastCellRow = 0;
            if (holdConsoleInput2 == null || holdConsoleInput2 == "") { secondaryLastCellRow = getLastRowWithValue(workbook, sheetInput2, secondaryFirstCellColumn); }
            else { secondaryLastCellRow = Convert.ToInt32(holdConsoleInput2) - 1; }

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

                Console.Write($"Arbeitsblatt (Nummer 1-32, default: {sheetInput1 + 1}): ");
                //resultSheet = Convert.ToInt32(Console.ReadLine()) - 1;
                resultSheet = getIntOrDefault(sheetInput1);

                Console.Write($"Ergebnis Spaltenwert (default: {getLastColumnWithValue(workbook, resultSheet, primaryFirstCellRow) + 2}): ");
                //resultColumn = Convert.ToInt32(Console.ReadLine()) - 1;
                resultColumn = getIntOrDefault(getLastColumnWithValue(workbook, resultSheet, primaryFirstCellRow) + 1);

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
            //printFittedSizeAsterixSurroundedText($"STARTING COMPUTATION WITH {varHold.getSimilarityMethodValue} ALGORITHM");
            //printFittedSizeAsterixSurroundedText($"STARTING COMPUTING WITH LEVENSHTEIN-DISTANCE ALGORITHM");
            //printFittedSizeAsterixSurroundedText($"COMPUTING SUPPORT VIA JACCARD-DISTANCE ALGORITHM");
            printFittedSizeAsterixSurroundedText("STARTING COMPUTATION");
            int matchedCells = 0;
            int matchedCellsIdentical = 0;
            if (verbose) { Console.WriteLine("starting stopwatch"); }
            stopwatch.Start();

            //Generating Checksum
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

                if (varHold.useDataFile)
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

                    if (varHold.useDataFile)
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
                        if (verbose == true)
                        {
                            Console.WriteLine(); printFittedSizeAsterixSurroundedText("MATCH--FOUND");
                        }
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
                        matchedCells++;
                        activeMatchValue = activeSecondaryCellValue;
                        activeMatchColumn = secondaryFirstCellColumn;
                        activeMatchRow = sCnt;
                        activeMatchPercentage = compareMatchPercentage;
                        activeMatchHammingDistance = compareMatchHammingDistance;
                    }
                    else if (compareMatchPercentage == activeMatchPercentage && compareMatchHammingDistance == activeMatchHammingDistance && compareMatchJaccardIndex < activeMatchJaccardIndex)
                    {
                        matchedCells++;
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


                    if (verbose == true)
                    {
                        if (compareMatch) { printFittedSizeAsterixRowCompact(); setConsoleBackgroundColorToGreen(); setConsoleColorToBlack(); }
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
                        if (varHold.useDataFile) { outputHold += ("ActivePrimaryCellValueOld: " + activePrimaryCellValueOld + " | "); }
                        outputHold += ("ActiveSecondaryCellValue: " + activeSecondaryCellValue + " | ");
                        if (varHold.useDataFile) { outputHold += ("ActiveSecondaryCellValueOld: " + activeSecondaryCellValueOld + " | "); }
                        outputHold += ("CompareIdentical: " + compareIdentical + " | ");
                        outputHold += ("CompareMatch: " + compareMatch + " | ");
                        outputHold += ("CompareMatchPercentage: " + compareMatchPercentage + " | ");
                        outputHold += ("CompareMatchHammingDistance: " + compareMatchHammingDistance + " | ");
                        outputHold += ("CompareMatchJaccardIndex: " + compareMatchJaccardIndex);

                        Console.WriteLine(outputHold);

                        if (compareMatch) { resetConsoleColor(); ; printFittedSizeAsterixRowCompact(); }
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
                    int columnHold = 0;

                    sheet = workbook.GetSheetAt(resultSheet);
                    IRow rrrow = sheet.GetRow(cnt);
                    ICell cccell = rrrow.CreateCell(resultColumn + (secondaryLastCellColumn - secondaryFirstCellColumn) + ++columnHold);
                    cellValueTransferHold = $"LD-Value: {activeMatchPercentage}";
                    cccell.SetCellValue(cellValueTransferHold);

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

                    if (varHold.useDataFile && (activePrimaryCellValueOld != null || activeSecondaryCellValueOld != null))
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
            }
            //Console.WriteLine("schreiben");
            Console.WriteLine();
            Console.WriteLine();
            if (writeResults) { printFittedSizeAsterixSurroundedText("SAVING"); }

            stopwatch.Stop();

            //if (writeResults) { workbook.Write(ffstream); }
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

            //stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            string timeSpanString = String.Format("{0:00}h:{1:00}m:{2:00}s:{3:00}ms", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds, timeSpan.Milliseconds);

            setConsoleColorToGreen();
            printFittedSizeAsterixSurroundedText("COMPUING--FINISHED");
            resetConsoleColor();

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("Matched Cells: " + matchedCells);
            Console.WriteLine("Matched Cells Identical: " + matchedCellsIdentical);
            Console.WriteLine();
            Console.Write("Old Checksum (SHA256): ");
            setConsoleColorToBlue();
            Console.WriteLine(varHold.checksumConvertedOriginal);
            resetConsoleColor();
            Console.Write("New Checksum (SHA256): ");
            setConsoleColorToBlue();
            Console.WriteLine(varHold.checksumConvertedNew);
            resetConsoleColor();
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

        private static int getIntOrDefault(int defaultValue)
        {
            string inputHold1 = Console.ReadLine();
            if (!string.IsNullOrEmpty(inputHold1)) { defaultValue = Convert.ToInt32(inputHold1) - 1; }

            return (defaultValue);
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
            VarHold varHold = new VarHold();

            int strLenght = text.Length;
            string outputHold = "";

            if (text.Contains("ERROR")) { setConsoleColorToRed(); }
            if (text.Contains("MATCH") && text.Contains("FOUND")) { setConsoleColorToGreen(); }

            printFittedSizeAsterixRowCompact();
            for (int cnt = 0; cnt <= (Console.WindowWidth / 2 - strLenght / 2) - 2; cnt++)
            {
                //Console.Write("*");
                outputHold += "*";
            }
            //Console.Write(" " + text + " ");
            outputHold += $" {text} ";
            int evenStr = 0;
            if (strLenght % 2 != 0) { evenStr = 1; }
            for (int cnt = 0; cnt <= (Console.WindowWidth / 2 - strLenght / 2) - 2 - evenStr; cnt++)
            {
                //Console.Write("*");
                outputHold += "*";
            }
            Console.WriteLine(outputHold);
            printFittedSizeAsterixRowCompact();
            resetConsoleColor();
            Console.WriteLine();
        }

        private static int getLastRowWithValue(IWorkbook workbook, int sheet, int column)
        {
            ISheet iSheet = workbook.GetSheetAt(sheet);

            int lastRowWithValue = 0;

            for (int i = iSheet.FirstRowNum; i <= iSheet.LastRowNum; i++)
            {
                IRow row = iSheet.GetRow(i);

                if (row != null)
                {
                    ICell cell = row.GetCell(column);

                    if (cell != null && cell.ToString() != null)
                    {
                        lastRowWithValue = i;
                    }
                }
            }

            return (lastRowWithValue);
        }

        private static int getLastColumnWithValue(IWorkbook workbook, int sheet, int row)
        {
            ISheet iSheet = workbook.GetSheetAt(sheet);

            int lastColumnWithValue = 0;

            IRow rrow = iSheet.GetRow(row);

            if (row != null)
            {
                for (int i = rrow.FirstCellNum; i <= rrow.LastCellNum; i++)
                {
                    ICell ccell = rrow.GetCell(i);

                    if (ccell != null && ccell.ToString() != "")
                    {
                        lastColumnWithValue = i;
                    }
                }
            }

            return (lastColumnWithValue);
        }

        private static void writeToSettingsFile(string path, string settingName, string settingValue)
        {
            VarHold varHold = new VarHold();
            Dictionary<string, string> settingsDict = new Dictionary<string, string>();

            if (File.Exists(path))
            {
                using (StreamReader file = new StreamReader(path))
                {
                    string line;
                    while ((line = file.ReadLine()) != null)
                    {
                        string[] streamParts = line.Split(' ');
                        if (streamParts.Length == 2)
                        {
                            settingsDict.Add(streamParts[0], streamParts[1]);
                        }
                    }
                }

                if (settingsDict.ContainsKey(settingName)) { settingsDict.Remove(settingName); }
            }

            settingsDict.Add(settingName, settingValue);

            using (StreamWriter file = new StreamWriter(path))
            {
                foreach (var entry in settingsDict)
                {
                    file.WriteLine("{0} {1}", entry.Key, entry.Value);
                }
            }
        }

        private static string readFromSettingsFile(string path, string settingName)
        {
            Dictionary<string, string> settingsDict = new Dictionary<string, string>();

            if (File.Exists(path))
            {
                using (StreamReader file = new StreamReader(path))
                {
                    string line;
                    while ((line = file.ReadLine()) != null)
                    {
                        string[] streamParts = line.Split(' ');
                        if (streamParts.Length == 2)
                        {
                            settingsDict.Add(streamParts[0], streamParts[1]);
                        }
                    }
                }
                if (settingsDict.ContainsKey(settingName)) { return settingsDict[settingName]; }
                else { return null; }
            }
            else { return null; }
        }

        private static void createDictionary(string jsonPath)
        {
            VarHold varHold = new VarHold();
            printFittedSizeAsterixSurroundedText("DATA FILE MODE");
            Console.WriteLine("No DATA file found");
        createDictionaryStart:
            Console.Write("Create DATA file (y/n): ");

            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    Console.WriteLine("y");
                    break;
                case ConsoleKey.N:
                    Console.Write("n");
                    Console.WriteLine();
                    Console.WriteLine();
                    run(false);
                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("DATA ERROR");
                    goto createDictionaryStart;
                    break;
            }
            Console.WriteLine();
            Console.WriteLine("### Eingabe der bekannten Übereinstimmungen ###");
            Console.WriteLine($"### quit: '{varHold.createDictionaryExitInput}' ###");

            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            int loopCnt = 0;

            while (true)
            {
            Start:
                Console.WriteLine();
                Console.WriteLine($"Datensatz-Index: {++loopCnt}");
                Console.Write("Eingabe primär: ");
                string primaryValue = Console.ReadLine();

                if (primaryValue.Contains(varHold.createDictionaryExitInput)) { goto CreateDictionarySave; }
                if (primaryValue.Contains(','))
                {
                    printFittedSizeAsterixSurroundedText("ERROR DATA");
                    goto Start;
                }
                if (dictionary.ContainsKey(primaryValue))
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Der folgende Wert existiert bereits und wurde nicht hinzugefügt: {primaryValue}");
                    Console.ResetColor();
                    goto Start;
                }

                Console.Write("Eingabe sekundär: ");
                string secondaryValue = Console.ReadLine();

                if (secondaryValue.Contains(varHold.createDictionaryExitInput)) { goto CreateDictionarySave; }
                if (secondaryValue.Contains(','))
                {
                    printFittedSizeAsterixSurroundedText("ERROR DATA");
                    goto Start;
                }
                if (dictionary.ContainsValue(secondaryValue))
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Der folgende Wert existiert bereits im Zusammenhang mit '{primaryValue}' und wurde nicht hinzugefügt: {secondaryValue}");
                    Console.ResetColor();
                    goto Start;
                }

                dictionary.Add(primaryValue, secondaryValue);
            }
        CreateDictionarySave:
            Console.WriteLine();
            Console.WriteLine("*** saving to file ***");
            Console.WriteLine();
            using (StreamWriter file = new StreamWriter(jsonPath))
            {
                foreach (var entry in dictionary)
                {
                    file.WriteLine("{0} {1}", entry.Key, entry.Value);
                }
            }
        }

        private static void createJson(string jsonPath)
        {
            VarHold varHold = new VarHold();
            printFittedSizeAsterixSurroundedText("JSON FILE MODE");
            Console.WriteLine("No JSON data file found");
        createJsonStart:
            Console.Write("Create JSON data file (y/n): ");

            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    Console.WriteLine("y");
                    break;
                case ConsoleKey.N:
                    Console.Write("n");
                    Console.WriteLine();
                    Console.WriteLine();
                    run(false);
                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine();
                    printFittedSizeAsterixSurroundedText("DATA ERROR");
                    goto createJsonStart;
                    break;
            }
            Console.WriteLine();
            Console.WriteLine("### Eingabe der bekannten Übereinstimmungen ###");
            Console.WriteLine($"### quit: '{varHold.createJsonExitInput}' ###");
            Console.WriteLine();
            int loopCnt = 0;
            Console.WriteLine($"Datensatz-Index: {++loopCnt}");
            Console.Write("Eingabe primär: ");
            string primaryValue = Console.ReadLine();
            Console.Write("Eingabe sekundär: ");
            string secondaryValue = Console.ReadLine();

            var valueHold = new createJsonValueHold
            {
                Value1 = primaryValue,
                Value2 = secondaryValue
            };

            var toJsonString = JsonSerializer.Serialize(valueHold);
            File.WriteAllText(jsonPath, toJsonString);

            while (true)
            {
                valueHold = JsonSerializer.Deserialize<createJsonValueHold>(toJsonString);

                Console.WriteLine();
                Console.WriteLine($"Datensatz-Index: {++loopCnt}");
                Console.Write("Eingabe primär: ");
                valueHold.Value1 = Console.ReadLine();
                Console.Write("Eingabe sekundär: ");
                valueHold.Value2 = Console.ReadLine();

                if (primaryValue.Contains(varHold.createJsonExitInput) || secondaryValue.Contains(varHold.createJsonExitInput)) { break; }

                toJsonString = JsonSerializer.Serialize(valueHold);
                File.WriteAllText(jsonPath, toJsonString);
            }
        }

        private static void setConsoleColorToRed() { Console.ForegroundColor = ConsoleColor.Red; }

        private static void setConsoleColorToGreen() { Console.ForegroundColor = ConsoleColor.Green; }

        private static void setConsoleColorToBlue() { Console.ForegroundColor = ConsoleColor.Blue; }

        private static void setConsoleColorToBlack() { Console.ForegroundColor = ConsoleColor.Black; }

        private static void setConsoleBackgroundColorToRed() { Console.BackgroundColor = ConsoleColor.Red; }

        private static void setConsoleBackgroundColorToGreen() { Console.BackgroundColor = ConsoleColor.Green; }

        private static void setConsoleBackgroundColorToBlue() { Console.BackgroundColor = ConsoleColor.Blue; }

        private static void resetConsoleColor() { Console.ResetColor(); }
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
        public double getJaccardIndex(string input1, string input2)
        {
            HashSet<char> hashSet1 = new HashSet<char>(input1);
            HashSet<char> hashSet2 = new HashSet<char>(input2);

            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }

            int intersectionValueCount = hashSet1.Intersect(hashSet2).Count();
            int unionValueCount = hashSet1.Union(hashSet2).Count();

            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }

            return ((double)intersectionValueCount / unionValueCount);
        }
        public int getHammingDistance(string stringHold1, string stringHold2)
        {
            prepareGetHammingDistance(ref stringHold1, ref stringHold2);
            int hammingDistance = 0;
            for (int i = 0; i < stringHold1.Length; i++)
            {
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                if (stringHold1[i] != stringHold2[i]) { hammingDistance++; }
            }

            return (hammingDistance);
        }
        private void prepareGetHammingDistance(ref string stringHold1, ref string stringHold2)
        {
            if (stringHold1.Length != stringHold2.Length)
            {
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                if (stringHold1.Length < stringHold2.Length)
                {
                    if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                    if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                    while (stringHold1.Length < stringHold2.Length)
                    {
                        if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                        if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                        stringHold1 += " ";
                    }
                }
                else
                {
                    while (stringHold1.Length > stringHold2.Length)
                    {
                        if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                        if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                        stringHold2 += " ";
                    }
                }

            }
        }
        private bool adaptStringLength(string stringHold1, string stringHold2)
        {
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            if (stringHold1.Length == stringHold2.Length) { return (false); }
            else if (stringHold1.Length < stringHold2.Length)
            {
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                while (stringHold1.Length < stringHold2.Length)
                {
                    if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                    if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                    stringHold1 += ""; //Prüfen
                }
            }
            else
            {
                while (stringHold1.Length > stringHold2.Length)
                {
                    if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                    if (toggleArrayAccessLog) { varhold.arrayAccess++; }
                    stringHold2 += "";
                }
            }

            return (true);
        }
        public short getStringContainingValue(string stringHold1, string stringHold2) //0: Equal; 1: 2 in 1; 2: 2 in 1; 3: no
        {
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
            if (toggleArrayAccessLog) { varhold.arrayAccess++; }
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
    public class createJsonValueHold
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }
    }
}