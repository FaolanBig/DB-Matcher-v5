using DB_Matching_main1;
using MathNet.Numerics.RootFinding;
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



            PrintIn.blue("DB-Matcher-v5 needs to be restarted");
            Program.shutdownOrRestart();
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
                    file.WriteLine($"{entry.Key} // {entry.Value}");
                }
            }
        }
    }
}
