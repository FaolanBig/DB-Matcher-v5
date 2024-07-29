using DB_Matching_main1;
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

                for (int i = 0; i < 10; i++)
                {
                    int position = 0;
                    int direction = 1;
                    int length = 5;
                    int threadSleepingDuration = 350;

                    Console.SetCursorPosition(0, Console.CursorTop);

                    for (int ii = 0; ii < length; ii++)
                    {
                        if (ii == position)
                        {
                            Console.Write("*");
                        }
                        else
                        {
                            Console.Write(" ");
                        }
                    }

                    position += direction;

                    if (position == length - 1 || position == 0)
                    {
                        direction *= -1;
                    }

                    Thread.Sleep(threadSleepingDuration);
                }
            }
        }
    }
}
