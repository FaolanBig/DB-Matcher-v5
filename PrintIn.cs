using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal class PrintIn
    {
        public static void blue(string toPrint)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine(toPrint);
            Console.ResetColor();
        }
        public static void red(string toPrint)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(toPrint);
            Console.ResetColor();
        }
        public static void green(string toPrint)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(toPrint);
            Console.ResetColor();
        }
        public static void yellow(string toPrint)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(toPrint);
            Console.ResetColor();
        }
    }
}
