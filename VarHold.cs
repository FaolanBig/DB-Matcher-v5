using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_Matcher_v5
{
    internal static class VarHold
    {
        //public const string repoURL = "https://github.com/FaolanBig/DB-Matcher-v5";
        public const string repoURLReportIssue = "https://github.com/FaolanBig/DB-Matcher-v5/issues/new";
        public const string repoURLReleases = "https://github.com/FaolanBig/DB-Matcher-v5/releases";
        public const string updateFileRepoUrl = "https://raw.githubusercontent.com/faolanbig/db-matcher-v5/master/version.txt";

        public static string currentMainFilePath;
        public static string getSimilarityMethodValue;
        /*public string GetSimilarityMethodValue
        {
            get { return getSimilarityMethodValue; }
            set { GetSimilarityMethodValue = value; }
        }*/
        //public BigInteger arrayAccess = BigInteger.Parse("0");
        public static long arrayAccess = 0;
        public static long cyclesLog = 0;
        public static string checksumConvertedOriginal = "NULL";
        public static string checksumConvertedNew = "NULL";
        public static int runHold1 = 0;
        public static int runHold2 = 0;
        public static string createJsonExitInput = "JSONEXIT";
        public static string createDictionaryExitInput = "DIREXIT";
        public static bool useDataFile = true;
        public static bool toggleConsoleColor = true;
        public static string currentSettingsFilePathHold = "";
        public static string currentRecoveryMenuFile = "";
        public static bool osIsWindows; //true: Windows; false: Linux
        public static string logFileNameInfo = "logfile.txt";
        public static string logFileNameError = logFileNameInfo;
        public static string logoFilePath = "";
        public static string currentHoldFilePath = "";
        public static string toPath = "";
        public static string excelPath = @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE";
        public const string dictionaryExcapeCharacterString = @"/\";
        public static string helperFilePath = "";
        public static string updateFilePath = "";
        public static string latestVersion = "";
        public static string currentVersion = "";

        public static int threadsQuantity = 0;
        public static double[] threads_progress;
        public static double[] threads_remainingTime;

        public static double thread1_progress = 0;
        public static double thread2_progress = 0;
        public static double thread3_progress = 0;
        public static double thread4_progress = 0;

        public static int thread1_progressPercentage = 0;
        public static int thread2_progressPercentage = 0;
        public static int thread3_progressPercentage = 0;
        public static int thread4_progressPercentage = 0;

        public static string thread1_remainingTime = "";
        public static string thread2_remainingTime = "";
        public static string thread3_remainingTime = "";
        public static string thread4_remainingTime = "";

        public static Dictionary<string, string> settings = new Dictionary<string, string>();

        public static int matchedCells = 0;
        public static int matchedCellsIdentical = 0;

        public static bool writeResults = false;
        public static int secondaryFromColumn;
        public static int secondaryToColumn;
    }
}
