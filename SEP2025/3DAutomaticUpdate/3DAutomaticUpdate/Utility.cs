using _3DAutomaticUpdate.model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate
{
    class Utility
    {
        public static String LTCCustomSheetName = "Input_Data";
        public static List<String> ModSheetsInSession = new List<string>();
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
                Utility.Log("ResetAlerts: " + ex.Message, logFilePath);
            }

        }

        public static void QuitSEEC(SolidEdge.Framework.Interop.Application objApp, String logFilePath)
        {
            try
            {
                Utility.Log("Quit Solid Edge..", logFilePath);
                objApp.Quit();
            }
            catch (Exception ex)
            {
                Utility.Log("QuitSEEC Exception.." + ex.Message, logFilePath);
                Utility.Log("QuitSEEC Exception.." + ex.StackTrace, logFilePath);
            }
        }
    }
}
