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
using NPOI.POIFS.Crypt;
using NPOI.SS.Formula.Functions;
using SixLabors.ImageSharp.Advanced;
using System;
using System.Collections;
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
            ToLog.Inf("SettingsAgent: @EditMode");
            Console.Clear();
            Program.printFittedSizeAsterixSurroundedText("Settings Agent");
            if (editManually)
            {
                if (File.Exists(VarHold.currentSettingsFilePathHold))
                {
                    ToLog.Inf("SettingsAgent: launching in external editor");
                    PrintIn.blue("launching in external editor");
                    if (VarHold.osIsWindows) { Process.Start(new ProcessStartInfo(VarHold.currentSettingsFilePathHold) { UseShellExecute = true }); }
                    if (!VarHold.osIsWindows) { Process.Start("nano", VarHold.currentSettingsFilePathHold); }

                    PrintIn.blue("DB-Matcher-v5 needs to be restarted");
                    Program.shutdownOrRestart();
                }
                else
                {
                    ToLog.Err("SettingsAgent: settings file not found");
                    PrintIn.red("settings file not found");
                    PrintIn.red("proceding with launcher");
                }
            }

            PrintIn.blue("starting settings launcher");
            PrintIn.blue("erasing current configuration");

            try
            {
                ToLog.Inf("SettingsAgent: erasing current configuration");
                VarHold.settings.Clear();
                PrintIn.green("erasing successful");
                ToLog.Inf("SettingsAgent: operation finished: success");
            }
            catch (Exception ex)
            {
                ToLog.Err($"SettingsAgent: an unexpected error occurred when erasing current configuration - error: {ex.Message}");
                PrintIn.red($"an unexpected error occurred: {ex.Message}");
                PrintIn.blue("proceeding");
            }

            Console.WriteLine();
            addYesNo("automatically use settings file instead of user dialog","automaticMode");
            addYesNo("use big logo at startUp", "useBigLogoAtStartUp");
            addYesNo("skip data file setup at startUp", "skipDataFileSetup");
            addYesNo("skip ask to use data file", "skipAskToUseDataFile");
            addYesNo("verbose output", "verbose");
            addYesNo("always write results", "writeResults");
            addYesNo("get an auditive feedback when finishing processes (beep)", "consoleBeep");
            addYesNo("visualize matching grade with color gradients", "colorGradient");
            addYesNo("enable multi threading (experimental)", "multiThreading");

            bool useAutoOpenExcel = false;
            if (VarHold.osIsWindows) { useAutoOpenExcel = addYesNo("automatically open newly generated excel-file in excel", "autoOpenExcel"); }
            if (useAutoOpenExcel && !File.Exists(VarHold.excelPath)) 
            {
                PrintIn.red($"file not found: excel.exe in {VarHold.excelPath}");
                if (RecoveryHandler.getYesOrNo("add a custom path to open excel files?")) { addValue("path to open excel files", "pathTo_excel.exe", VarHold.excelPath); }

            }
            else if (useAutoOpenExcel) { PrintIn.blue($"using excel.exe in {VarHold.excelPath}"); }

            Console.WriteLine();
            PrintIn.yellow("changing the settings will overwrite any existing setting file");
            Program.setConsoleColorToYellow();

            AskToSave:
            Console.Write("continue and save new settings? (y/n): ");
            string userInput = Console.ReadLine();

            switch (userInput)
            {
                case "y":
                    break;
                case "n":
                    PrintIn.red("abort");
                    PrintIn.red("returning to recovery menu");
                    PrintIn.WigglyStarInBorders(runs: 1);
                    RecoveryHandler.RunRecovery();
                    break;
                default:
                    goto AskToSave;
            }


            try
            {
                ToLog.Inf("SettingsAgent: saving settings");
                PrintIn.blue("saving settings");
                PrintIn.WigglyStarInBorders(runs: 1);
                SaveSettings();
                PrintIn.green("saving successful");
                ToLog.Inf("SettingsAgent: operation finished: success");
            }
            catch (Exception ex)
            {
                ToLog.Err($"SettingsAgent: an unexpected error occurred when saving settings - error: {ex.Message}");
                PrintIn.red($"an unexpected error occurred: {ex.Message}");
                PrintIn.red($"if this is a bug, please report it on {VarHold.repoURLReleases}");
            }
            
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
                    //Console.WriteLine("y");
                    ToLog.Inf($"SettingsAgent: setting added: {key}{VarHold.dictionaryExcapeCharacterString}{valueYes}");
                    VarHold.settings.Add(key, valueYes);
                    PrintIn.green($"added: {key}{VarHold.dictionaryExcapeCharacterString}{valueYes}");
                    return true;
                case "n":
                    ToLog.Inf($"SettingsAgent: setting added: {key}{VarHold.dictionaryExcapeCharacterString}{valueYes}");
                    VarHold.settings.Add(key, valueNo);
                    PrintIn.green($"added: {key}{VarHold.dictionaryExcapeCharacterString}{valueNo}");
                    //Console.WriteLine("n");
                    return false;
                default:
                    PrintIn.red("bad input");
                    goto Start;
            }
        }
        internal static string addValue(string toAsk, string key, string defaultValue = "")
        {
            START:
            if (!string.IsNullOrEmpty(defaultValue)) 
            {
                Console.Write(toAsk + $" - enter value (default: {defaultValue}): "); 
            }
            else { Console.Write(toAsk + " - enter value: "); }
            string userInput = Console.ReadLine();

            if (string.IsNullOrEmpty(userInput))
            {
                if (!string.IsNullOrEmpty(defaultValue)) { userInput = defaultValue; }
                else { PrintIn.red("bad input"); goto START; }
            }
            //PrintIn.yellow($"confirm input: {key}//{userInput}");
            bool getYesOrNo = RecoveryHandler.getYesOrNo($"confirm input: {key}{VarHold.dictionaryExcapeCharacterString}{userInput}");
            if (getYesOrNo) 
            {
                ToLog.Inf($"SettingsAgent: setting added: {key}{VarHold.dictionaryExcapeCharacterString}{userInput}");
                VarHold.settings.Add(key, userInput);
                PrintIn.green($"added: {key}{VarHold.dictionaryExcapeCharacterString}{userInput}"); 
                return userInput; 
            }
            else 
            { 
                PrintIn.red("abort, try again");
                goto START; 
            }
        }
        internal static void FileLookUp()
        {
            if (!File.Exists(VarHold.currentSettingsFilePathHold))
            {
                ToLog.Err("SettingsAgent: error: settings file not found");
                PrintIn.red("no settings file found");
                PrintIn.blue("launching recovery mode");
                PrintIn.WigglyStarInBorders();
                RecoveryHandler.RunRecovery();
            }
            LoadSettings();
        }
        internal static void LoadSettings()
        {
            ToLog.Inf("SettingsAgent: loading settings");
            try
            {
                foreach (var line in File.ReadLines(VarHold.currentSettingsFilePathHold))
                {
                    var parts = line.Split(new[] { $"{VarHold.dictionaryExcapeCharacterString}" }, StringSplitOptions.None);
                    if (parts.Length == 2)
                    {
                        VarHold.settings[parts[0].Trim()] = parts[1].Trim();
                    }
                }
                ToLog.Inf("SettingsAgent: operation finished: success");
            }
            catch (Exception ex)
            {
                ToLog.Err($"SettingsAgent: an unexpected error occurred when loading settings - error: {ex.Message}");
                PrintIn.red($"an unexpected error occured - see log for more information");
                PrintIn.red($"please report it to {VarHold.repoURLReleases}");
            }
        }
        internal static void SaveSettings()
        {
            ToLog.Inf("SettingsAgent: saving settings");
            try
            {
                using (StreamWriter file = new StreamWriter(VarHold.currentSettingsFilePathHold))
                {
                    foreach (var entry in VarHold.settings)
                    {
                        file.WriteLine($"{entry.Key}{VarHold.dictionaryExcapeCharacterString}{entry.Value}");
                    }
                }
                ToLog.Inf("SettingsAgent: operation finished: success");
            }
            catch (Exception ex)
            {
                ToLog.Err($"SettingsAgent: an unexpected error occurred when saving settings - error: {ex.Message}");
                PrintIn.red($"an unexpected error occured - see log for more information");
                PrintIn.red($"please report it to {VarHold.repoURLReleases}");
            }
        }
        internal static void ViewSettings()
        {
            ToLog.Inf("SettingsAgent: viewing settings");
            Console.Clear();
            Program.printFittedSizeAsterixSurroundedText("Settings Agent");

            PrintIn.blue("file lookup");

            if (!File.Exists(VarHold.currentSettingsFilePathHold))
            {
                ToLog.Err("SettingsAgent: no settings file found");
                PrintIn.red("no settings file found");
                PrintIn.blue("try adding a settings configuration in recovery mode");
                //PrintIn.blue("returning to recovery menu");
                //PrintIn.WigglyStarInBorders();
                RecoveryHandler.WaitForKeystrokeENTER("hit ENTER to return to recovery menu");
                RecoveryHandler.RunRecovery();
            }
            else
            {
                PrintIn.green("settings file found");
                PrintIn.blue("loading settings");
                LoadSettings();

                Console.WriteLine();
                //Console.WriteLine("⤓⤓⤓");
                foreach (var i in VarHold.settings)
                {
                    //Console.WriteLine($">>> {i.Key}//{i.Value}");
                    if (i.Value == "true") { Console.Write($">>> {i.Key}{VarHold.dictionaryExcapeCharacterString}"); PrintIn.green(i.Value); }
                    else if (i.Value == "false") { Console.Write($">>> {i.Key}{VarHold.dictionaryExcapeCharacterString}"); PrintIn.red(i.Value); }
                    else { Console.Write($">>> {i.Key}{VarHold.dictionaryExcapeCharacterString}"); PrintIn.yellow(i.Value); }
                    Thread.Sleep(10);
                }
                //Console.WriteLine("⤒⤒⤒");
                Console.WriteLine();
                RecoveryHandler.WaitForKeystrokeENTER("hit ENTER to return to recovery menu");
                RecoveryHandler.RunRecovery();
            }
        }
        internal static string GetSettingValue(string keyToCheck)
        {
            VarHold.settings.TryGetValue(keyToCheck, out string value);
            if (string.IsNullOrEmpty(value))
            {
                ToLog.Err($"SettingsAgent: value is null or empty - key: {keyToCheck} @GetSettingValue()");
                PrintIn.red($"error when reading key-value (value is null or empty) of key: {keyToCheck}");
                PrintIn.red("this could be caused by a currupt settings file");
                PrintIn.red("to solve this, reconfigure the settings in recovery mode");
                PrintIn.red($"if this is a bug, please report it on {VarHold.repoURLReleases}");
                RecoveryHandler.WaitForKeystrokeENTER();
                PrintIn.blue("launching recovery mode");
                PrintIn.WigglyStarInBorders();
                RecoveryHandler.RunRecovery();
                return null;
            }
            else { return value; }
        }
        internal static bool GetSettingIsTrue(string keyToCheck)
        {
            if (GetSettingValue(keyToCheck) == "true") { return true; } else { return false; }
        }
    }
}