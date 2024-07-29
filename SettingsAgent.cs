using DB_Matching_main1;
using MathNet.Numerics.RootFinding;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal class SettingsAgent
    {
        public static void EditMode()
        {
        }
        internal static void FileLookUp()
        {
            if (!File.Exists(VarHold.currentSettingsFilePathHold))
            {
                PrintIn.red("no settings file found");
                PrintIn.red("launching recovery mode");
                Console.WriteLine();

                PrintIn.wigglyStarInBorders();
                
                RecoveryHandler.RunRecovery();
            }
        }
    }
}
