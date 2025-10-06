using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidEdge.StructureEditor.Interop;

namespace DemoAddInTC.utils
{
    class SE_SESSION
    {
        static SolidEdgeFramework.Application objApp = null;
        static SolidEdge.RevisionManager.Interop.Application objReviseApp = null;
        static RevisionManager.Application objReviseApp1 = null;
        static SolidEdge.StructureEditor.Interop.Application StructureEditorApplication = null;

        public static void setStructureEditorSession(SolidEdge.StructureEditor.Interop.Application oApp)
        {
            StructureEditorApplication = oApp;
        }

        public static SolidEdge.StructureEditor.Interop.Application getStructureEditorSession()
        {
            return StructureEditorApplication;
        }

        public static void InitializeStructureEditorSession(String logFilePath)
        {
            try
            {
                Utlity.Log("SE_SESSION - Creating StructureEditor Instance ... ", logFilePath);
                StructureEditorApplication = (SolidEdge.StructureEditor.Interop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("StructureEditor.Application"), true);
                        
                Utlity.Log("SE_SESSION - Created...StructureEditor Instance .", logFilePath);

            }
            catch (Exception ex)
            {
                Utlity.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utlity.Log(ex.StackTrace.ToString(), logFilePath);
                return;
            }
        }

        public static void QuitStructureEditor()
        {
            StructureEditorApplication.Quit();
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
                objReviseApp.DisplayAlerts = 0;
                Utlity.Log("SE_SESSION - Created...Revision Manager Instance .", logFilePath);

            }
            catch (Exception ex)
            {
                Utlity.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utlity.Log(ex.StackTrace.ToString(), logFilePath);
                return;
            }
        }

        public static void InitializeSolidEdgeRevisionManagerSession1(String logFilePath)
        {
            try
            {
                Utlity.Log("SE_SESSION - Creating Revision Manager Instance1 ... ", logFilePath);
                objReviseApp1 = new RevisionManager.Application();
                objReviseApp1.DisplayAlerts = 0;
                Utlity.Log("SE_SESSION - Created...Revision Manager Instance1 .", logFilePath);

            }
            catch (Exception ex)
            {
                Utlity.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utlity.Log(ex.StackTrace.ToString(), logFilePath);
                return;
            }
        }

        
        public static RevisionManager.Application getRevisionManagerSession1()
        {
            return objReviseApp1;
        }

        public static void killRevisionManager1(String logFilePath)
        {
            objReviseApp1.Quit();
            objReviseApp1 = null;
        }

        public static SolidEdge.RevisionManager.Interop.Application getRevisionManagerSession()
        {
            return objReviseApp;
        }

        public static void killRevisionManager(String logFilePath)
        {
            objReviseApp.DisplayAlerts = 1;
            objReviseApp.Quit();
            objReviseApp = null;
        }
        
    }
}
