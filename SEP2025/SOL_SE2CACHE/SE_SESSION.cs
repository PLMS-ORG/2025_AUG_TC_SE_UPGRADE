using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidEdge.Framework.Interop;
using System.Threading;

namespace LTC_SE2CACHE
{
    class SE_SESSION
    {
        static SolidEdge.Framework.Interop.Application objApp = null;

        static SolidEdgeFramework.Application app = null;
        static int MaxEdgeInstanceTryOutTimes = 3;
        static int currentCreateProcessIndex = 1;
        static double WaitTime = 5.00;


        public static void InitializeSolidEdgeSession1()
        {
            try
            {
                //Get Active session of Solid Edge 
                Console.WriteLine("SE_SESSION - Initiating Solid Edge Application");
                // 24 - OCT, Connecting to an Existing Instance is Risk. Because the Other Process can Quit at Any Time
                //app = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, false);
                app = SolidEdgeCommunity.SolidEdgeUtils.Start();
            }
            catch (Exception ex)
            {
                if (currentCreateProcessIndex < MaxEdgeInstanceTryOutTimes)
                {
                    Console.WriteLine("SE_SESSION - Wait for " + WaitTime.ToString() + " Minutes");
                    Thread.Sleep(TimeSpan.FromMinutes(WaitTime));
                    currentCreateProcessIndex++;
                    InitializeSolidEdgeSession1();
                }

                Console.WriteLine("SE_SESSION - " + ex.Message.ToString());
                Console.WriteLine(ex.StackTrace.ToString());
                return;
            }
        }

        public static SolidEdgeFramework.Application getSolidEdgeSession1()
        {
            return app;
        }

        public static void killSolidEdgeSession1()
        {
            try
            {

                if (app != null)
                {
                    Console.WriteLine("SE_SESSION - Killing Solid Edge Application");
                    //objApp.Visible = true;                    
                    //if (app.Documents != null && app.Documents.Count > 0) app.Documents.Close();
                    //app.DoIdle();
                    app.Quit();

                    // Releasing Object after Quiting Causing the Problem ?? No Idea. 13 OCT 2018
                    Marshal.FinalReleaseComObject(app);
                    app = null;
                    //
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SE_SESSION - " + ex.Message.ToString());
                Console.WriteLine(ex.StackTrace.ToString());
                return;

            }

        }

        public static void InitializeSolidEdgeSession(String logFilePath)
        {

            try
            {
                //Get Active session of Solid Edge 
                Utility.Log("SE_SESSION - Initating Solid Edge Application", logFilePath);
                Process[] pname = Process.GetProcessesByName("Edge");
                if (pname.Length != 0)
                {
                    Utility.Log("SE_SESSION - Identified Edge EXE : " + pname[0].ProcessName, logFilePath);
                    objApp = (SolidEdge.Framework.Interop.Application)Marshal.GetActiveObject("SolidEdge.Application");
                }
                else
                {
                    Utility.Log("SE_SESSION - Creating Edge Instance ... ", logFilePath);
                    objApp = (SolidEdge.Framework.Interop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"), true);
                    objApp = (SolidEdge.Framework.Interop.Application)Marshal.GetActiveObject("SolidEdge.Application");
                }
            }
            catch (Exception ex)
            {
                Utility.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utility.Log(ex.StackTrace.ToString(), logFilePath);
                return;
            }
        }
        

        public static SolidEdge.Framework.Interop.Application getSolidEdgeSession()
        {
            return objApp;
        }


        public static void killSolidEdgeSession(String logFilePath)
        {
            try
            {

                if (objApp != null)
                {

                    Utility.Log("SE_SESSION - Killing Solid Edge Application", logFilePath);
                    //objApp.Visible = true;                    
                    if (objApp.Documents != null && objApp.Documents.Count > 0) objApp.Documents.Close();
                    objApp.DoIdle();
                    objApp.Quit();

                    // Releasing Object after Quiting Causing the Problem ?? No Idea. 13 OCT 2018
                    //Marshal.ReleaseComObject(objApp);
                    //objApp = null;
                    //
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ex)
            {
                Utility.Log("SE_SESSION - " + ex.Message.ToString(), logFilePath);
                Utility.Log(ex.StackTrace.ToString(), logFilePath);
                return;

            }

        }
    }
}
