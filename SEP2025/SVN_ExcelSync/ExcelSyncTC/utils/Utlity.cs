using ExcelSyncTC.model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSyncTC.utils
{
    class Utlity
    {
        // 29 - SEPT - Change for Skipping LTC Custom Sheets during Sync & other Functionalities
        public static String LTCCustomSheetName = "Input_Data";
        public static List<String> ModSheetsInSession = new List<string>();

        public static void ResetAlerts(SolidEdgeFramework.Application objApp, bool AlertFlag,String logFilePath)
        {
            try
            {
                if (objApp != null)
                {
                    objApp.DisplayAlerts = AlertFlag;
                }

            }
            catch (Exception ex)
            {
                utils.Utlity.Log("ResetAlerts: " + ex.Message, logFilePath);
            }

        }

        public static String CreateLogDirectory()
        {
            String HomeFolder = "";
            String DE = "3DE_LTC";
            String creoHome = "";
            HomeFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (HomeFolder != null && HomeFolder.Equals("") == false)
            {
                creoHome = Path.Combine(HomeFolder, DE);
                if (Directory.Exists(creoHome) == false)
                {
                    Directory.CreateDirectory(creoHome);
                }
            }
            return creoHome;
        }

        public static void Log(string logMessage, string logFilePath,[Optional]String Option)
        {
            try
            {
                if (Option == null || Option.Equals("") == true)
                {
                    //DateTime currentDateTime = DateTime.Now;
                    string formattedDateTime = DateTime.Now.ToString("yyyy-MM-dd || HH:mm:ss");
                    StreamWriter w = File.AppendText(logFilePath);
                    w.WriteLine("{0}", formattedDateTime + ":::" + logMessage);
                    Console.WriteLine(logMessage);
                    w.Close();
                }
                else if (Option.Equals("INFO", StringComparison.OrdinalIgnoreCase) == true)
                {
                    LogToForm(logMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Writing Exception: " + ex.Message);
            }

        }

        private static void LogToForm(string logMessage)
        {
            Debug.WriteLine(logMessage);

        }

        private static void LogToEdge(string logMessage)
        {
            SolidEdgeFramework.Application app = SE_SESSION.getSolidEdgeSession();
            app.StatusBar = logMessage;

        }

        // Contains latest FeatureLine Data from UI.
        public static Dictionary<String, List<FeatureLine>> BuildFeatureDictionary(List<FeatureLine> allFeatureLinesList, String logFilePath)
        {

            Dictionary<String, List<FeatureLine>> featureDictionary = new Dictionary<String, List<FeatureLine>>();

            foreach (FeatureLine varr in allFeatureLinesList)
            {
                //Log(varr.PartName, logFilePath);
                if (featureDictionary.ContainsKey(varr.PartName) == true)
                {
                    //Log("Fetching Entry FOR: " + varr.PartName, logFilePath);
                    List<FeatureLine> flList = null;
                    featureDictionary.TryGetValue(varr.PartName, out flList);
                    flList.Add(varr);
                }
                else
                {
                    //Log("First Entry FOR: " + varr.PartName, logFilePath);
                    List<FeatureLine> flList = new List<FeatureLine>();
                    flList.Add(varr);
                    featureDictionary.Add(varr.PartName, flList);
                }
            }

            return featureDictionary;

        }

        // Contains latest Variable Data from UI.
        public static Dictionary<String, List<Variable>> BuildVariableDictionary(List<Variable> AllVariablesList, String logFilePath)
        {

            Dictionary<String, List<Variable>> variablesDict = new Dictionary<String, List<Variable>>();

            foreach (Variable varr in AllVariablesList)
            {
                //Log(varr.PartName + ":::" + varr.name + ":::" + varr.systemName, logFilePath);

                if (variablesDict.ContainsKey(varr.PartName) == true)
                {
                    //Log("Fetching Entry FOR: " + varr.PartName, logFilePath);
                    List<Variable> variablesList = null;
                    variablesDict.TryGetValue(varr.PartName, out variablesList);
                    variablesList.Add(varr);
                }
                else
                {
                    //Log("First Entry FOR: " + varr.PartName, logFilePath);
                    List<Variable> variablesList = new List<Variable>();
                    variablesList.Add(varr);
                    variablesDict.Add(varr.PartName, variablesList);
                }
            }

            return variablesDict;

        }

        public static void HideExcelColumn(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, String columnIndex, String logFilePath)
        {
            //Log("HIDING COLUMN " + columnIndex,logFilePath);
            // Hide the Column Entirely
            if (xlWorksheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
            {
                //Utlity.Log("Sheet Visible: " + xlWorksheet.Name , logFilePath);
            }
            else
            {

                //Utlity.Log("Sheet InVisible: " + xlWorksheet.Name , logFilePath);
                xlWorksheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;
            }
            Microsoft.Office.Interop.Excel.Range delRange1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.get_Range(columnIndex + ":" + columnIndex, Missing.Value);
            delRange1.EntireColumn.Select();
            delRange1.EntireColumn.Hidden = true;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(delRange1);

        }

        public static void HideExcelRow(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, int RowIndex, String logFilePath)
        {
           
            xlWorksheet.Activate();
            var hiddenRange = xlWorksheet.Range[xlWorksheet.Cells[RowIndex, 1], xlWorksheet.Cells[RowIndex, 1]];
            hiddenRange.EntireRow.Hidden = true;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(hiddenRange);
            hiddenRange = null;

        }

        public static bool parseLog(String logFilePath, [Optional] String startString)
        {
            if (startString == null) startString = "";

            String stageDir = Utlity.CreateLogDirectory();
            String parseLog_LogFilePath = System.IO.Path.Combine(stageDir, "parseLog_LogFile" + ".txt");

            // Get the config file Path from the executable directory
            //string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            //Log("exeDirectory: " + exeDirectory, parseLog_LogFilePath);
            //string configFilePath = Path.Combine(exeDirectory, "error_config.txt");


            // Check if config file and log file exist
            //if (!File.Exists(configFilePath))
            //{
            //    //Console.WriteLine("Config file not found: " + configFilePath);
            //    Log("Config file not found: " + configFilePath, parseLog_LogFilePath);
            //    return false;
            //}
            if (!File.Exists(logFilePath))
            {
                //Console.WriteLine("Log file not found: " + logFilePath);
                Log("Log file not found: " + logFilePath, parseLog_LogFilePath);
                return false;
            }

            // Read search keywords from config file (one per line)
            //List<string> keywords = new List<string>(File.ReadAllLines(configFilePath));

            // keywords in the list of strings will be hard coded below.
            List<string> keywords = new List<string>();
            keywords.Add("exception");
            keywords.Add("error");
            keywords.Add("failed");
            keywords.Add("unable");
            keywords.Add("hresult");
            keywords.Add("warning");
            
          
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
                Log("Total lines in log file: " + allLines.Length, parseLog_LogFilePath);
                Log("Searching for start string: " + startString, parseLog_LogFilePath);
                //int startIndex = Array.LastIndexOf(allLines, allLines.FirstOrDefault(line => line.Contains(startString)));
                int startIndex = Array.FindLastIndex(allLines, line => line.Contains(startString));
                Log("Start index found at: " + startIndex, parseLog_LogFilePath);
                if (startIndex != -1)
                {
                    allLines = allLines.Skip(startIndex).ToArray();

                }
                Log("Lines to be parsed after applying start string: " + allLines.Length, parseLog_LogFilePath);
            }
            else
            {
                allLines = File.ReadAllLines(logFilePath);
            }

            if (allLines.Length == 0)
            {
                Log("No lines found in the log file after the specified start string.", parseLog_LogFilePath);
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

            Log(Path.GetFileName(logFilePath) + " is Parsed Successfully", parseLog_LogFilePath);
            //Log("Keywords searched from config file: " + configFilePath, parseLog_LogFilePath);
            // Display the output
            foreach (var kvp in keywordCounts)
            {
                //Console.WriteLine($"{kvp.Key}: {kvp.Value}");
                Log($"{kvp.Key}: {kvp.Value}", parseLog_LogFilePath);
            }

            Log("keywordCounts.Count: " + keywordCounts.Count, parseLog_LogFilePath);
            if (keywordCounts.Values.Sum() == 0)
            {
                //Console.WriteLine("No keywords found in the log file.");
                Log("No keywords found in the log file.", parseLog_LogFilePath);
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
