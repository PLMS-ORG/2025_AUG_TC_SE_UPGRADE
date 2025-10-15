using DemoAddInTC.se;
using DemoAddInTC.services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DemoAddInTC.utils
{
    class Utility
    {
        public static String DE = "creo_data_preparation";

        public static bool downloadFolderExists(String sdExpFile, String logFilePath)
        {

            String stageDir = System.IO.Path.GetDirectoryName(sdExpFile);
            String folderName_sdExp = System.IO.Path.GetFileNameWithoutExtension(sdExpFile);
            String creoHome = Path.Combine(stageDir, folderName_sdExp);
            if (Directory.Exists(creoHome) == false)
            {
                return false;
            }
            else
            {
                Utility.Log(sdExpFile + " Exists Already..", logFilePath);
                return true;
            }

        }

        public static bool CreateDownloadFolderMM(String sdExpFile, String logFilePath)
        {
            String stageDir = System.IO.Path.GetDirectoryName(sdExpFile);
            String folderName_sdExp = System.IO.Path.GetFileNameWithoutExtension(sdExpFile);

            {
                String creoHome = Path.Combine(stageDir, folderName_sdExp);
                if (Directory.Exists(creoHome) == false)
                {
                    Directory.CreateDirectory(creoHome);
                }
                else
                {
                    try
                    {
                        // delete
                        Directory.Delete(creoHome);
                        Directory.CreateDirectory(creoHome);
                    }
                    catch (Exception ex)
                    {
                        Log("Exception in Delete: " + ex.Message, logFilePath);
                        return false;
                    }
                }


                CreoUtilitySession.DownloadFolder = creoHome;
                //Log("CreoUtilitySession.DownloadFolder: " + CreoUtilitySession.DownloadFolder, logFilePath);
                return true;
            }
        }

        public static void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }

            Directory.Delete(target_dir, false);
        }


        public static String CreateDownloadFolder(String stageDir, String logFilePath, String itemID)
        {

            {
                String creoHome = Path.Combine(stageDir, itemID);
                if (Directory.Exists(creoHome) == false)
                {
                    Directory.CreateDirectory(creoHome);
                }
                else
                {
                    try
                    {

                        DeleteDirectory(creoHome);
                        Directory.CreateDirectory(creoHome);
                    }
                    catch (Exception ex)
                    {
                        Log("Exception in Delete: " + ex.Message, logFilePath);
                        return "";
                    }
                }


                CreoUtilitySession.DownloadFolder = creoHome;
                //Log("CreoUtilitySession.DownloadFolder: " + CreoUtilitySession.DownloadFolder, logFilePath);
                return creoHome;
            }
        }

        public static void Log(string logMessage, string logFilePath)
        {
            try
            {
                StreamWriter w = File.AppendText(logFilePath);
                w.WriteLine("{0}", logMessage);
                Console.WriteLine(logMessage);
                w.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Writing Exception: " + ex.Message);
            }

        }

        public static bool parseLog(String logFilePath, [Optional] String startString)
        {
            if (startString == null) startString = "";

            String stageDir = Utlity.CreateLogDirectory();
            String parseLog_LogFilePath = System.IO.Path.Combine(stageDir, "parseLog_LogFile" + ".txt");

            // Get the config file Path from the executable directory
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string configFilePath = Path.Combine(exeDirectory, "error_config.txt");

            // Check if config file and log file exist
            if (!File.Exists(configFilePath))
            {
                //Console.WriteLine("Config file not found: " + configFilePath);
                Utility.Log("Config file not found: " + configFilePath, parseLog_LogFilePath);
                return false;
            }
            if (!File.Exists(logFilePath))
            {
                //Console.WriteLine("Log file not found: " + logFilePath);
                Utility.Log("Log file not found: " + logFilePath, parseLog_LogFilePath);
                return false;
            }

            // Read search keywords from config file (one per line)
            List<string> keywords = new List<string>(File.ReadAllLines(configFilePath));

            // Dictionary to store keyword counts
            Dictionary<string, int> keywordCounts = new Dictionary<string, int>();
            foreach (var keyword in keywords)
            {
                keywordCounts[keyword] = 0;
            }

            var allLines = new string[] { };
            // if StartString is provided, search for the last occurrence and start parsing from there
            if (!string.IsNullOrEmpty(startString))
            {
                allLines = File.ReadAllLines(logFilePath);
                int startIndex = Array.LastIndexOf(allLines, allLines.FirstOrDefault(line => line.Contains(startString)));
                if (startIndex != -1)
                {
                    allLines = allLines.Skip(startIndex).ToArray();

                }
            }
            else
            {
                allLines = File.ReadAllLines(logFilePath);
            }

            if (allLines.Length == 0)
            {
                Utility.Log("No lines found in the log file after the specified start string.", parseLog_LogFilePath);
                return false;
            }

            // Read the log file line by line
            foreach (var line in allLines)
            {
                foreach (var keyword in keywords)
                {
                    int count = CountOccurrences(line, keyword);
                    keywordCounts[keyword] += count;
                }
            }

            Log(logFilePath + " Parsed Successfully", parseLog_LogFilePath);
            Log("Keywords searched from config file: " + configFilePath, parseLog_LogFilePath);
            // Display the output
            foreach (var kvp in keywordCounts)
            {
                //Console.WriteLine($"{kvp.Key}: {kvp.Value}");
                Utility.Log($"{kvp.Key}: {kvp.Value}", parseLog_LogFilePath);
            }

            Utility.Log("keywordCounts.Count: " + keywordCounts.Count, parseLog_LogFilePath);
            if (keywordCounts.Values.Sum() == 0)
            {
                //Console.WriteLine("No keywords found in the log file.");
                Utility.Log("No keywords found in the log file.", parseLog_LogFilePath);
                return true;
            }
            return false; // keywords were found, indicating potential issues
        }

        // Helper function to count occurrences of a substring in a string (case-insensitive)
        static int CountOccurrences(string source, string substring)
        {
            int count = 0, index = 0;
            while ((index = source.IndexOf(substring, index, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                count++;
                index += substring.Length;
            }
            return count;
        }

    }
}
