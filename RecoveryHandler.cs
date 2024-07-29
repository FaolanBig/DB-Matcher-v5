using DB_Matching_main1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal class RecoveryHandler
    {
        public static void StartUp()
        {
            if (Console.KeyAvailable)
            {
                if (Console.ReadKey(true).Key == ConsoleKey.Escape)
                {
                    Console.Clear();
                    Program.printFittedSizeAsterixSurroundedText("Recovery");

                    SettingsAgent.EditMode(VarHold.currentSettingsFilePathHold);

                    Console.WriteLine("Press ENTER to continue");
                    while (Console.ReadKey(true).Key != ConsoleKey.Enter) { }

                    Console.Clear();
                }
            }
        }
    }
}
