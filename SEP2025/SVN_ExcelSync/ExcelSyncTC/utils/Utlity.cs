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
    }
}
