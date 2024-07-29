/*
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


using DB_Matching_main1;
using MathNet.Numerics.RootFinding;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal class SettingsAgent
    {
        public static void EditMode(bool editManually = false)
        {
            Console.Clear();
            Program.printFittedSizeAsterixSurroundedText("Settings Agent");
            if (editManually)
            {
                if (File.Exists(VarHold.currentSettingsFilePathHold))
                {
                    PrintIn.blue("launching in external editor");
                    if (VarHold.osIsWindows) { Process.Start(new ProcessStartInfo(VarHold.currentSettingsFilePathHold) { UseShellExecute = true }); }
                    if (!VarHold.osIsWindows) { Process.Start("nano", VarHold.currentSettingsFilePathHold); }

                    PrintIn.blue("DB-Matcher-v5 needs to be restarted");
                    Program.shutdownOrRestart();
                }
                else
                {
                    PrintIn.red("settings file not found");
                    PrintIn.red("proceding with launcher");
                }
            }

            PrintIn.blue("starting settings launcher");
            Console.WriteLine();
            addYesNo("automatically use settings file instead of user dialog","automaticMode");

            SaveSettings();
            PrintIn.blue("DB-Matcher-v5 needs to be restarted");
            Program.shutdownOrRestart();
        }
        internal static bool addYesNo(string toAsk, string key, string valueYes = "true", string valueNo = "false")
        {
            Start:
            Console.Write(toAsk + " (y/n): ");
            string userInput = Console.ReadLine();

            switch (userInput)
            {
                case "y":
                    Console.WriteLine("y");
                    VarHold.settings.Add(key, valueYes);
                    PrintIn.green($"added: {key}//{valueYes}");
                    return true;
                case "n":
                    VarHold.settings.Add(key, valueNo);
                    PrintIn.green($"added: {key}//{valueNo}");
                    Console.WriteLine("n");
                    return false;
                default:
                    PrintIn.red("bad input");
                    goto Start;
            }
        }
        internal static void FileLookUp()
        {
            if (!File.Exists(VarHold.currentSettingsFilePathHold))
            {
                PrintIn.red("no settings file found");
                PrintIn.blue("launching recovery mode");
                Console.WriteLine();

                PrintIn.wigglyStarInBorders();
                
                RecoveryHandler.RunRecovery();
            }
            loadSettings();
        }
        internal static void loadSettings()
        {
            foreach (var line in File.ReadLines(VarHold.currentSettingsFilePathHold))
            {
                var parts = line.Split(new[] { "//" }, StringSplitOptions.None);
                if (parts.Length == 2)
                {
                    VarHold.settings[parts[0].Trim()] = parts[1].Trim();
                }
            }
        }
        internal static void SaveSettings()
        {
            using (StreamWriter file = new StreamWriter(VarHold.currentSettingsFilePathHold))
            {
                foreach (var entry in VarHold.settings)
                {
                    file.WriteLine($"{entry.Key}//{entry.Value}");
                }
            }
        }
    }
}
