using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _2DAutomaticUpdate
{
    class Utility
    {
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
               Log("ResetAlerts: " + ex.Message, logFilePath);
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



        public static void QuitSEEC(SolidEdgeFramework.Application objApp, String logFilePath)
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
