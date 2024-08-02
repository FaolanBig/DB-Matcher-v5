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
            try
            {
                var task = GetVersionHandler();
                task.GetAwaiter().GetResult();
                Console.WriteLine($"version: {VarHold.latestVersion}");
                Console.ReadLine();
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
                    ToLog.Inf($"pulling latest version number from {VarHold.updateFileRepoUrl}");
                    HttpResponseMessage response = await client.GetAsync(VarHold.updateFileRepoUrl);
                    response.EnsureSuccessStatusCode();
                    VarHold.latestVersion = await response.Content.ReadAsStringAsync();
                }
                catch (Exception ex)
                {
                    ToLog.Err($"an error occurred during version check - error: {ex.Message}");
                    PrintIn.red("an error occurred during version check");
                    PrintIn.red("see log for more information");
                    VarHold.latestVersion = "";
                    RecoveryHandler.WaitForKeystrokeENTER();
                }
            }
        }
    }
}
