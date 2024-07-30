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
                .WriteTo.Console()
                .WriteTo.File(VarHold.logFileNameInfo)
                .CreateLogger();

            Serilog.Log.Information(toLog);
            Serilog.Log.CloseAndFlush();
        }
        public static void Err(string toLog)
        {
            Serilog.Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File(VarHold.logFileNameError)
                .CreateLogger();

            Serilog.Log.Error(toLog);
            Serilog.Log.CloseAndFlush();
        }
    }
}
