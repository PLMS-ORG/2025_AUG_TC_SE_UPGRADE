using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSyncTC.utils
{
    class SE_SESSION
    {
        static SolidEdgeFramework.Application objApp = null;
        static SolidEdge.RevisionManager.Interop.Application objReviseApp = null;

        public static void InitializeSolidEdgeSession(String logFilePath)
        {
            try
            {
                //Get Active session of Solid Edge 
                Utlity.Log("SE_SESSION - Initating Solid Edge Application", logFilePath);
                Process[] pname = Process.GetProcessesByName("Edge");
                if (pname.Length != 0)
                {
                    Utlity.Log("SE_SESSION - Identified Edge EXE : " + pname[0].ProcessName, logFilePath);
                    objApp = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                    //objApp.Documents.Close();
                   // objApp.DoIdle();
                }
                else
                {
                    Utlity.Log("SE_SESSION - Creating Edge Instance ... ", logFilePath);
                    objApp = (SolidEdgeFramework.Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"), true);
                    objApp = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                    objApp.Documents.Close();
                    objApp.DoIdle();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utlity.Log(ex.StackTrace.ToString(), logFilePath);
                return;
            }
        }

        public static void setSolidEdgeSession(SolidEdgeFramework.Application oApp)
        {
            objApp = oApp;
        }

        public static SolidEdgeFramework.Application getSolidEdgeSession()
        {
            return objApp;
        }

        public static void InitializeSolidEdgeRevisionManagerSession(String logFilePath)
        {
            try
            {
                Utlity.Log("SE_SESSION - Creating Revision Manager Instance ... ", logFilePath);
                objReviseApp = new SolidEdge.RevisionManager.Interop.Application();

            }
            catch (Exception ex)
            {
                Utlity.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utlity.Log(ex.StackTrace.ToString(), logFilePath);
                return;
            }
        }

        public static SolidEdge.RevisionManager.Interop.Application getRevisionManagerSession()
        {
            return objReviseApp;
        }

        public static void killRevisionManager(String logFilePath)
        {
            objReviseApp.Quit();
        }
        
    }
}
