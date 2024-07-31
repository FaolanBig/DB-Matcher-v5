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


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DB_Matching_main1;
using NPOI.SS.Formula.Functions;
using Serilog;

namespace DB_Matcher_v5
{
    internal class Safety
    {
    }
    internal static class ToLog
    {
        public static void Inf(string toLog)
        {
            Serilog.Log.Logger = new LoggerConfiguration()
                //.WriteTo.Console()
                .WriteTo.File(VarHold.logFileNameInfo)
                .CreateLogger();

            Serilog.Log.Information(toLog);
            Serilog.Log.CloseAndFlush();
        }
        public static void Err(string toLog)
        {
            Serilog.Log.Logger = new LoggerConfiguration()
                //.WriteTo.Console()
                .WriteTo.File(VarHold.logFileNameError)
                .CreateLogger();

            Serilog.Log.Error(toLog);
            Serilog.Log.CloseAndFlush();
        }
    }
}
