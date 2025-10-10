using Log;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;


namespace DemoAddInTC.utils
{
    class Utlity
    {
        public static String LTCCustomSheetName = "Input_Data";

        public static void Log(string logMessage, string logFilePath)
        {
            if (logFilePath == null || logFilePath.Equals("") == true)
            {
                Console.WriteLine("Please provide valid logFile Path");
                return;
            }
            try
            {
                Console.WriteLine(logMessage);
                if(!File.Exists(logFilePath)) 
                {
                    File.Create(logFilePath).Close();               
                }
                StreamWriter w = File.AppendText(logFilePath);
                w.WriteLine("{0}", logMessage);
                w.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Writing Exception: " + ex.Message);
            }

        }
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
                    Utlity.Log("ResetAlerts: " + ex.Message, logFilePath);

                }
            }

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
                Utlity.Log("NO FILES FOUND IN "+ folderToPublish,logFilePath);
                return false;
            }
            String SeFile = Path.Combine(folderToPublish, SeFiles[0]);
            if (System.IO.File.Exists(SeFile) == true)
            {
                Utlity.Log("FILES FOUND IN " + folderToPublish, logFilePath);
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

   

        private static void LogToForm(string logMessage)
        {
            //Debug.WriteLine(logMessage);
            Trace.WriteLine(logMessage);

        }

        public static String getLTCCustomSheetName()
        {
            //MessageBox.Show(PropertyFile);
            String PropertyFile = Utlity.getPropertyFilePath();
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                log.write(logType.ERROR,"Could Not Find Property File " + PropertyFile);
                return "";
            }
            Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                log.write(logType.ERROR,"Cannot Parse Property File to Get Templates Folder: " + PropertyFile);
                return "";
            }
            String customSheetName = Props.get("CUSTOM_SHEET_NAME");
            if (customSheetName == null)
            {
                log.write(logType.ERROR, "CUSTOM_SHEET_NAME is not Set." + customSheetName);
                return "";
            }
            return customSheetName;
        }

        public static String getPropertyValue(String PropertyKey)
        {
            String PropertyFile = Utlity.getPropertyFilePath();
            String tc_mode = "";
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                log.write(logType.ERROR, "getManageMode: Could Not Find Property File " + PropertyFile);
                return "";
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                log.write(logType.ERROR, "Cannot Parse Property File to Get Templates Folder: " + PropertyFile);
                return "";
            }
            tc_mode = Props.get(PropertyKey);
            if (tc_mode == null)
            {
                log.write(logType.ERROR, PropertyKey + " is not Set." + tc_mode);
                return "";
            }

            return tc_mode;
        }
    }
}
