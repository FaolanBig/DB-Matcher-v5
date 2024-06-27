/*
    !!!This program will install DB-Matcher on your system by building it from source files!!!

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


using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading;

namespace INSTALLER
{
    public static class VarHold
    {
        public static string repositoryURL = "https://github.com/FaolanBig/DB-Matcher-v5";
        //ypublic static string repositoryName = @"DB-Matcher-v5";
        public static string repositoryName = @"";
        public static string currentFilePath;
        public static string currentDirPath;
        public static string toDirPath;
        public static bool autonomousInstall = false;
        public static bool dotnetAvailable = false;
        public static bool gitAvailable = false;
    }
    internal class Program
    {
        private static void Main(string[] args)
        {
        Start:
            Console.Clear();
            Console.WriteLine("this installer will install DB-Matcher and its dependencies");

            if (!UserInteractions.switchYesNo()) { exitProgramm(); }

            if (UserInteractions.switchYesNo("Proceed as autonomous installation? (y/n): ")) { VarHold.autonomousInstall = true; }

            GetTargetPath:
            Console.Write("target directory (this if empty): ");
            string userDefinedDirHold = Console.ReadLine();
            /*if (!string.IsNullOrEmpty(userDefinedDirHold))
            {
                if (!Directory.Exists(userDefinedDirHold))
                {
                    printInRed($"invalid path: {userDefinedDirHold}");
                    goto GetTargetPath;
                }
                userDefinedDirHold = userDefinedDirHold.Replace("\"", "");
                VarHold.currentDirPath = AppDomain.CurrentDomain.BaseDirectory;
            }
            else
            {
                VarHold.currentDirPath = userDefinedDirHold;
            }*/

            if (string.IsNullOrEmpty(userDefinedDirHold))
            {
                VarHold.currentDirPath = AppDomain.CurrentDomain.BaseDirectory;
            }
            else
            {
                if (Directory.Exists(userDefinedDirHold))
                {
                    userDefinedDirHold = userDefinedDirHold.Replace("\"", "");
                    VarHold.currentDirPath = userDefinedDirHold;
                }
            }

            Console.Write("using path: ");
            printInBlue(VarHold.currentDirPath);

            //check for commands availability
            printInBlue("starting installation");
            printInBlue("command availablity scan");

            VarHold.dotnetAvailable = ApiInteractions.CheckIfCMDCommandIsAvailableForExecute("dotnet --version");
            VarHold.gitAvailable = ApiInteractions.CheckIfCMDCommandIsAvailableForExecute("git help");

            //installation of non-available commandsA

            if (!VarHold.gitAvailable)
            {
                Console.Write("installing: ");
                printInBlue("git");
                if (!VarHold.autonomousInstall && !UserInteractions.switchYesNo()) { exitProgramm(); }

                ApiInteractions.ExecuteInCMD("curl -O -L https://github.com/git-for-windows/git/releases/download/v2.35.1.windows.2/Git-2.35.1.2-64-bit.exe && Git-2.35.1.2-64-bit.exe");
                restartProgram();
            }

            if (!VarHold.dotnetAvailable)
            {
                Console.Write("install: ");
                printInBlue("dotnet");
                if (!VarHold.autonomousInstall && !UserInteractions.switchYesNo()) { exitProgramm(); }

                ApiInteractions.ExecuteInCMD("git clone https://github.com/dotnet/core.git\r\n && cd core/eng && powershell -ExecutionPolicy Bypass -File install-dotnet.ps1");
                restartProgram();
            }

            //installation DB-Matcher
            printInGreen("all commands available");
            Console.Write("installing: ");
            printInBlue("DB-Matcher");

            Console.WriteLine("get directory");
            VarHold.currentDirPath += @"DB-Matcher_git-clone";

            Console.WriteLine("find empty directory");
            string dirHold = VarHold.currentDirPath;
            int cnt = 1;

            while (Directory.Exists(dirHold))
            {
                dirHold = VarHold.currentDirPath + cnt;
                cnt++;
            }

            VarHold.toDirPath = dirHold;

            Console.Write("new directory: ");
            printInBlue(VarHold.toDirPath);

            Console.Write("cloning repository from: ");
            printInBlue(VarHold.repositoryURL);
            ApiInteractions.ExecuteInCMD($"git clone {VarHold.repositoryURL} {VarHold.toDirPath}");

            Console.WriteLine("updating directory");
            VarHold.toDirPath += VarHold.repositoryName;

            Console.WriteLine("installing packages");
            ApiInteractions.ExecuteInCMD($"cd {VarHold.toDirPath} && dotnet add package NPOI");

            Console.WriteLine("installing dependencies");
            ApiInteractions.ExecuteInCMD($"cd {VarHold.toDirPath} && dotnet restore");

            Console.WriteLine("building from source-code");
            ApiInteractions.ExecuteInCMD($"cd {VarHold.toDirPath} && dotnet build");

            printInGreen("installation finished");
            Console.WriteLine("press any key to exit the installer");
            Console.ReadKey();
        }
        public static void colorReset() { Console.ResetColor(); }
        public static void green() { Console.ForegroundColor = ConsoleColor.Green; }
        public static void red() { Console.ForegroundColor = ConsoleColor.Red; }
        public static void blue() { Console.ForegroundColor = ConsoleColor.Blue; }
        public static void printInGreen(string toPrint, bool writeLine = true)
        {
            green();
            if (writeLine) { Console.WriteLine(toPrint); }
            else { Console.Write(toPrint); }
            colorReset();
        }
        public static void printInRed(string toPrint, bool writeLine = true)
        {
            red();
            if (writeLine) { Console.WriteLine(toPrint); }
            else { Console.Write(toPrint); }
            colorReset();
        }
        public static void printInBlue(string toPrint, bool writeLine = true)
        {
            blue();
            if (writeLine) { Console.WriteLine(toPrint); }
            else { Console.Write(toPrint); }
            colorReset();
        }
        public static void exitProgramm()
        {
            printInRed("exit");
            Thread.Sleep(1000);
            Environment.Exit(0);
        }
        public static void restartProgram()
        {
            printInRed("restarting");

            var currentFileName = Assembly.GetExecutingAssembly().Location;
            Process.Start(currentFileName);

            exitProgramm();
        }
        public static void restartComputer()
        {
            ApiInteractions.ExecuteInCMD("shutdown /r /t 5 /c \"restarting from INSTALLER.exe (DB-Matcher) in 5sek\"");
        }
        public static void askToRestartOrExit()
        {
            printInBlue("your computer has to be restarted");
            printInRed("save any open files, then proceed");
            if (UserInteractions.switchYesNo("restart computer? (y/n): ")) { restartComputer(); }
            exitProgramm();
        }
    }

    internal static class UserInteractions
    {
        public static bool switchYesNo(string displayedSwitchInfo = "proceed? (y/n): ") //1: yes; 0: no
        {
        Start:
            Console.Write(displayedSwitchInfo);

            switch (Console.ReadKey(true).Key)
            {
                case ConsoleKey.Y:
                    Program.green();
                    Console.WriteLine("y");
                    Program.colorReset();
                    return (true);
                    break;
                case ConsoleKey.N:
                    Program.red();
                    Console.WriteLine("n");
                    Program.colorReset();
                    return (false);
                    break;
                default:
                    Console.WriteLine();
                    goto Start;
            }
        }
    }

    internal static class ApiInteractions
    {
        public static void ExecuteInCMD(string cmd)
        {
            Console.Write("using shell: ");
            Program.printInBlue("cmd.exe");
            Console.Write("executing in new window: ");
            Program.printInBlue(cmd);

            /*string cmdHold = "/C" + cmd;
            System.Diagnostics.Process.Start("CMD.exe", cmdHold);*/

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/C {cmd}",
                //RedirectStandardOutput = true,
                //RedirectStandardError = true,
                UseShellExecute = true,
                //CreateNoWindow = true
            };

            Process process = new Process
            {
                StartInfo = startInfo
            };

            Console.WriteLine("attempting to run command");
            process.Start();

            //string output = process.StandardOutput.ReadToEnd();
            //string error = process.StandardError.ReadToEnd();

            Console.WriteLine("waiting for process to finish");

            process.WaitForExit();

            Program.printInGreen("process finished");

            int exitCode = process.ExitCode;

            /*if (exitCode != 0)
            {
                Program.printInRed($"error while running command: {cmd}");
                Program.printInRed($"exit code: {exitCode}");
                GetErrorMessage(cmd);
                if (!UserInteractions.switchYesNo("continue? (y/n): ")) { Program.exitProgramm(); }
            }*/
        }
        public static string GetErrorMessage(string cmd) //1: available; 0: not available
        {
            Console.Write("using shell: ");
            Program.printInBlue("cmd.exe");
            Console.Write($"fetching error message: ");
            Program.printInBlue(cmd);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/C {cmd}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                //CreateNoWindow = true
            };

            Process process = new Process
            {
                StartInfo = startInfo
            };

            Console.WriteLine("attempting to run command");
            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            Console.WriteLine("waiting for process to finish");

            process.WaitForExit();

            Console.WriteLine("process finished");

            int exitCode = process.ExitCode;

            /*if (exitCode != 0 || !string.IsNullOrEmpty(error)) //not available (error occured)
            {
                Program.printInRed("command not available");
                return false;
            }
            else //available
            {
                Program.printInGreen("command available");
                return true;
            }*/
            if (string.IsNullOrEmpty(error))
            {
                Program.printInRed("no error message has been fetched");
                Program.printInBlue($"error code: {exitCode}");
                return "null";
            }
            else
            {
                Program.printInRed($"error message: {error}");
                Program.printInBlue($"{exitCode}");
                return error;
            }
        }
        public static bool CheckIfCMDCommandIsAvailableForExecute(string cmd) //1: available; 0: not available
        {
            Console.Write("using shell: ");
            Program.printInBlue("cmd.exe");
            Console.Write($"checking for command: ");
            Program.printInBlue(cmd);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/C {cmd}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                //CreateNoWindow = true
            };

            Process process = new Process
            {
                StartInfo = startInfo
            };

            Console.WriteLine("attempting to run command");
            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            Console.WriteLine("waiting for process to finish");

            process.WaitForExit();

            Console.WriteLine("process finished");

            int exitCode = process.ExitCode;

            if (exitCode != 0 || !string.IsNullOrEmpty(error)) //not available (error occured)
            {
                Program.printInRed("command not available");
                return false;
            }
            else //available
            {
                Program.printInGreen("command available");
                return true;
            }
        }

    }
}
