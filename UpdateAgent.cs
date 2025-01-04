using DB_Matching_main1;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DB_Matcher_v5
{
    internal static class UpdateAgent
    {
        internal static string GetVersionLatest()
        {
            ToLog.Inf("UpdateAgent: running host process");
            try
            {
                var task = GetVersionHandler();
                task.GetAwaiter().GetResult();
                return VarHold.latestVersion;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"error: {ex.Message}");
                return null;
            }
        }
        private static async Task GetVersionHandler()
        {
            using (HttpClient client = new())
            {
                try
                {
                    ToLog.Inf($"UpdateAgent: pulling latest version number from {VarHold.updateFileRepoUrl}");
                    HttpResponseMessage response = await client.GetAsync(VarHold.updateFileRepoUrl);
                    response.EnsureSuccessStatusCode();
                    VarHold.latestVersion = await response.Content.ReadAsStringAsync();
                }
                catch (Exception ex)
                {
                    ToLog.Err($"UpdateAgent: an error occurred during version check - error: {ex.Message}");
                    PrintIn.red("an error occurred during version check");
                    PrintIn.red("see log for more information");
                    PrintIn.blue("check your internet connection");
                    VarHold.latestVersion = "";
                    RecoveryHandler.WaitForKeystrokeENTER();
                }
            }
        }
        internal static bool CheckForUpdates() //0: currently running the latest version of DB-Matcher-v5; 1: update available
        {
            GetVersionLatest();

            try
            {
                ToLog.Inf("getting current version");
                VarHold.currentVersion = File.ReadAllText(VarHold.updateFilePath);
                ToLog.Inf($"process finished: success - current version: {VarHold.currentVersion}");
            }
            catch (Exception ex)
            {
                ToLog.Err($"an error occurred when reading file: {VarHold.updateFilePath} - error: {ex.Message}");

                PrintIn.red($"an error occurred when reading file: {VarHold.updateFilePath}");
                PrintIn.red($"try downloading the latest release of DB-Matcher-v5 on {VarHold.repoURLReleases}");
                PrintIn.red("see log for more information");
            }

            VarHold.currentVersion = VarHold.currentVersion.TrimEnd('\r', '\n');
            VarHold.latestVersion = VarHold.latestVersion.TrimEnd('\r', '\n');

            if (string.IsNullOrEmpty(VarHold.latestVersion)) { return false; }
            else if (VarHold.currentVersion != VarHold.latestVersion)
            {
                ToLog.Inf($"UpdateAgent: new version available on {VarHold.repoURLReleases} (current version: {VarHold.currentVersion} --> latest version: {VarHold.latestVersion})");
                PrintIn.blue($"new version awailable on {VarHold.repoURLReleases}");
                PrintIn.blue($"current version: {VarHold.currentVersion} --> latest version: {VarHold.latestVersion}");
                RecoveryHandler.WaitForKeystrokeENTER();
                return true;
            }
            else if (VarHold.currentVersion == VarHold.latestVersion)
            {
                ToLog.Inf($"UpdateAgent: currently using the latest version of DB-Matcher-v5 (version: {VarHold.currentVersion})");
                PrintIn.green($"currenly running the latest version of DB-Matcher-v5: {VarHold.currentVersion}");
                return false;
            }
            ToLog.Err("UpdateAgent: unexpected error at version check occurred");
            return false;
        }
    }
}
