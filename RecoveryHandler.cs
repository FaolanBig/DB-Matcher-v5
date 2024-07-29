using DB_Matching_main1;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal class RecoveryHandler
    {
        public static void StartUp()
        {
            PrintIn.blue("press ESC to enter recovery mode");

            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < 1000)
            {
                if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape)
                {
                    PrintIn.blue("starting recovery mode...");
                    Console.WriteLine();
                    PrintIn.wigglyStarInBorders();

                    RunRecovery();
                    return;
                }

                Thread.Sleep(50);
            }

        }
        public static void RunRecovery()
        {
            Console.Clear();
            Program.printFittedSizeAsterixSurroundedText("Recovery");

            Menu();

            PrintIn.blue("DB-Matcher-v5 needs to be restarted");
            Program.shutdownOrRestart();
        }
        internal static bool Menu()
        {
            string content;
            Console.WriteLine("launching menu");
            if (!File.Exists(VarHold.currentRecoveryMenuFile))
            {
                PrintIn.red("no menu file found");
                PrintIn.red("to fix this problem, follow the following steps");
                Console.WriteLine();
                Console.WriteLine($"   1. download \"recoveryMenu.txt\" from {VarHold.repoURL}");
                Console.WriteLine($"   2. move \"recoveryMenu.txt\" to {VarHold.currentRecoveryMenuFile}");
                Console.WriteLine("   3. restart DB-Matcher-v5");
                Console.WriteLine();
                PrintIn.blue("this should fix the problem");
                PrintIn.blue($"if this is a bug, please report it by creating an issue on {VarHold.repoURL}");
                WaitForKeystrokeENTER();
                return true;
            }
            else
            {
                try
                {
                    Console.WriteLine($"loading menu from {VarHold.currentRecoveryMenuFile}");
                    content = File.ReadAllText(VarHold.currentRecoveryMenuFile);

                    if (string.IsNullOrEmpty(content))
                    {
                        PrintIn.red($"an unexpected error occurred when reading {VarHold.currentRecoveryMenuFile}: content is null or empty");

                        PrintIn.red("to fix this problem, follow the following steps");
                        Console.WriteLine();
                        Console.WriteLine($"   1. download \"recoveryMenu.txt\" from {VarHold.repoURL}");
                        Console.WriteLine($"   2. replace {VarHold.currentRecoveryMenuFile} with the newly downloaded file");
                        Console.WriteLine("    3. restart DB-Matcher-v5");
                        Console.WriteLine();
                        PrintIn.blue("this should fix the problem");
                        PrintIn.blue($"if this is a bug, please report it by creating an issue on {VarHold.repoURL}");
                        WaitForKeystrokeENTER();
                        return false;
                    }
                }
                catch ( Exception ex )
                {
                    PrintIn.red($"an unexpected error occurred when reading {VarHold.currentRecoveryMenuFile}");
                    PrintIn.red($"error message: {ex.Message}");

                    PrintIn.red("to fix this problem, follow the following steps");
                    Console.WriteLine();
                    Console.WriteLine($"   1. download \"recoveryMenu.txt\" from {VarHold.repoURL}");
                    Console.WriteLine($"   2. replace {VarHold.currentRecoveryMenuFile} with the newly downloaded file");
                    Console.WriteLine("    3. restart DB-Matcher-v5");
                    Console.WriteLine();
                    PrintIn.blue("this should fix the problem");
                    PrintIn.blue($"if this is a bug, please report it by creating an issue on {VarHold.repoURL}");
                    WaitForKeystrokeENTER();
                    return true;
                }
                PrintIn.green("loading finished");
                Console.WriteLine();
                Console.WriteLine(content);
                Console.WriteLine();
                EnterNumber:
                Console.Write("enter number: ");
                string userInput = Console.ReadLine();

                bool isNumber = int.TryParse(userInput, out int number);

                if (string.IsNullOrEmpty(userInput))
                {
                    PrintIn.red("bad input");
                    goto EnterNumber;
                }
                if (!isNumber)
                {
                    PrintIn.red("bad input");
                    goto EnterNumber;
                }

                switch (userInput)
                {
                    case "1":
                        PrintIn.blue("launching SettingsAgent");
                        SettingsAgent.EditMode();
                        break;
                    default:
                        PrintIn.red("bad input");
                        goto EnterNumber;
                }
                
                return false;
            }
        }
        public static void WaitForKeystrokeENTER(string outputHold = "Press ENTER to continue")
        {
            PrintIn.blue("Press ENTER to continue");
            while (Console.ReadKey(true).Key != ConsoleKey.Enter) { }
        }
    }
}
