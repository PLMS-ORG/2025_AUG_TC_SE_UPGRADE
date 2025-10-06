using DemoAddInTC.model;
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

namespace DemoAddInTC.utils
{
    class Utlity
    {
        public static String LTCCustomSheetName = "Input_Data";

        public static void ResetAlerts(SolidEdgeFramework.Application objApp, bool AlertFlag, String logFilePath)
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
                if (logFilePath != null && logFilePath.Equals("") == false)
                {
                    utils.Utlity.Log("ResetAlerts: " + ex.Message, logFilePath);

                }
            }

        }

        public static String getLTCCustomSheetName(String logFilePath)
        {
            //MessageBox.Show(PropertyFile);
            String PropertyFile = Utlity.getPropertyFilePath();
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                Log("Could Not Find Property File " + PropertyFile, logFilePath);
                return "";
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                Log("Cannot Parse Property File to Get Templates Folder: " + PropertyFile,logFilePath);
                return "";
            }
            String customSheetName = Props.get("CUSTOM_SHEET_NAME");
            if (customSheetName == null)
            {
                Log("CUSTOM_SHEET_NAME is not Set." + customSheetName,logFilePath);
                return "";
            }
            return customSheetName;
        }
        public static String getPropertyFilePath()
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;
            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            String PropertyFile = Path.Combine(ClickOnceLocation, "Property.txt");
            return PropertyFile;

        }

        public static String getManageMode(String logFilePath)
        {
            String PropertyFile = Utlity.getPropertyFilePath();
            String tc_mode = "";
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                Log("getManageMode: Could Not Find Property File " + PropertyFile, logFilePath);
                return "";
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                Log("Cannot Parse Property File to Get Templates Folder: " + PropertyFile,logFilePath);
                return "";
            }
            tc_mode = Props.get("TC_MODE");
            if (tc_mode == null)
            {
                Log("TC_MODE is not Set." + tc_mode,logFilePath);
                return "";
            }

            return tc_mode;
        }

        public static String getPropertyValue(String PropertyKey, String logFilePath)
        {
            String PropertyFile = Utlity.getPropertyFilePath();
            String tc_mode = "";
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                Log("getManageMode: Could Not Find Property File " + PropertyFile, logFilePath);
                return "";
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                Log("Cannot Parse Property File to Get Templates Folder: " + PropertyFile, logFilePath);
                return "";
            }
            tc_mode = Props.get(PropertyKey);
            if (tc_mode == null)
            {
                Log(PropertyKey + " is not Set." + tc_mode, logFilePath);
                return "";
            }

            return tc_mode;
        }
        

        public static String getExecutingPath()
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;
            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            return ClickOnceLocation;

        }

        public static bool checkifFilesAlreadyInFolderToPublish(String folderToPublish,String logFilePath)
        {

            string[] SeFiles = Directory.GetFiles(folderToPublish, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.EndsWith(".par") || x.EndsWith(".asm") || x.EndsWith(".psm")))
                                         .ToArray();
            if (SeFiles == null || SeFiles.Length == 0)
            {
                Log("NO FILES FOUND IN "+ folderToPublish,logFilePath);
                return false;
            }
            String SeFile = Path.Combine(folderToPublish, SeFiles[0]);
            if (System.IO.File.Exists(SeFile) == true)
            {
                Log("FILES FOUND IN " + folderToPublish, logFilePath);
                return true;
            }


            return false;
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


        public static String ConvertFromDBVal<T>(object obj)
        {
            if (obj == null || obj == System.DBNull.Value)
            {
                return ""; // returns the default value for the type
            }
            else
            {
                return (String)obj.ToString();
            }
        }

        public static void Log(string logMessage, string logFilePath,[Optional]String Option)
        {
            if (logFilePath == null || logFilePath.Equals("") == true)
            {
                return;
            }
            try
            {
                if (Option == null || Option.Equals("") == true)
                {
                    StreamWriter w = File.AppendText(logFilePath);
                    w.WriteLine("{0}", logMessage);
                    Console.WriteLine(logMessage);
                    w.Close();
                    LogToEdge(logMessage);
                }
                else if (Option.Equals("INFO", StringComparison.OrdinalIgnoreCase) == true)
                {
                    LogToForm(logMessage);
                }
                else if (Option.Equals("CTD", StringComparison.OrdinalIgnoreCase) == true)
                {
                    StreamWriter w = File.AppendText(logFilePath);
                    w.WriteLine("{0}", logMessage);
                    Console.WriteLine(logMessage);
                    w.Close();

                    try
                    {
                        StringBuilder sb = new StringBuilder(Ribbon2d.tcw.statusTextBox.Text);
                        sb.AppendLine(logMessage);
                        Ribbon2d.tcw.statusTextBox.Text = sb.ToString(); //("\n" + logMessage);
                    }
                    catch (Exception)
                    {
                        
                    }
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Writing Exception: " + ex.Message);
            }

        }

        private static void LogToForm(string logMessage)
        {
            //Debug.WriteLine(logMessage);
            Trace.WriteLine(logMessage);

        }

        private static void LogToEdge(string logMessage)
        {
            SolidEdgeFramework.Application app = SE_SESSION.getSolidEdgeSession();
            app.StatusBar = logMessage;
            
        }

        public static DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
           
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }

        public static void HideExcelColumn(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet,String columnIndex,String logFilePath)
        {
            //Log("HIDING COLUMN " + columnIndex,logFilePath);
            // Hide the Column Entirely
            if (xlWorksheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
            {
                //Utlity.Log("Sheet Visible: " + xlWorksheet.Name , logFilePath);
            }else {

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
            //Log("HIDING COLUMN " + columnIndex,logFilePath);
            // Hide the Column Entirely
            //Microsoft.Office.Interop.Excel.Range delRange1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.get_Range("A" + ":" + RowIndex.ToString(), Missing.Value);
            //xlWorksheet.Activate();
            //xlWorksheet.Cells[RowIndex, 1].EntireRow.Select();
            //xlWorksheet.Cells[RowIndex, 1].EntireRow.Hidden = true;
            //delRange1.EntireRow.Select();
            //delRange1.EntireRow.Hidden = true;
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(delRange1);

            //Microsoft.Office.Interop.Excel.Range delRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.get_Range((Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 1], (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 1]);
            //Microsoft.Office.Interop.Excel.Range delRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.get_Range("A" + ":" + RowIndex.ToString(), Missing.Value);
            //delRange.EntireRow.Select();
            //delRange.EntireRow.Hidden = true;         
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(delRange);
            xlWorksheet.Activate();
            var hiddenRange = xlWorksheet.Range[xlWorksheet.Cells[RowIndex, 1], xlWorksheet.Cells[RowIndex, 1]];
            hiddenRange.EntireRow.Hidden = true;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(hiddenRange);
            hiddenRange = null;

        }

        public static void AutoFitExcelColumn(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, String columnIndex, String logFilePath)
        {
            //Log("AUTOFIT COLUMN " + columnIndex, logFilePath);
            // Hide the Column Entirely
            //Microsoft.Office.Interop.Excel.Range delRange1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.get_Range(columnIndex + "1", Missing.Value);           
            xlWorksheet.Columns[columnIndex + ":" + columnIndex].EntireColumn.AutoFit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(delRange1);

        }


        // Contains latest Variable Data from UI.
        public static Dictionary<String, List<Variable>> BuildVariableDictionary(List<Variable> AllVariablesList,String logFilePath)
        {

            Dictionary<String, List<Variable>> variablesDict = new Dictionary<String, List<Variable>>();

            foreach (Variable varr in AllVariablesList)
            {
                //Log(varr.PartName + ":::" + varr.name + ":::" + varr.systemName, logFilePath);
                
                if (variablesDict.ContainsKey(varr.PartName) == true)
                {
                    //Log("Fetching Entry FOR: " + varr.PartName, logFilePath);
                    List<Variable>variablesList = null;
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

        public static String getValue(DataRow row, String propName, String logFilePath)
        {
            String value = "";
            try
            {
                value = (String)row[propName];
            }
            catch (Exception ex)
            {
                //Utlity.Log(ex.Message + ":::" + propName, logFilePath);
                return "";
            }
            return value;
        }

        public static bool getBoolValue(DataRow row, String propName, String logFilePath)
        {
            bool value = false;
            try
            {
                //value = (bool)row[propName];
                Boolean sValue = (Boolean)row[propName];
                value = sValue;            
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message + ":::" + propName, logFilePath);
                return false;
            }
            return value;
        }

        
    }
}
