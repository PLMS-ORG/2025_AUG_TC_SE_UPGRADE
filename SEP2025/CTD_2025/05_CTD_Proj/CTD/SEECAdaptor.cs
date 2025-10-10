using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DemoAddInTC.utils;
using SolidEdgeConstants;
using SolidEdge.StructureEditor.Interop;
using Log;

namespace DemoAddInTC.se
{
    public class SEECAdaptor
    {
        static String userName = "";
        static String password ="";
        static String group = "";
        static String Role = "";
        public static String URL = ""; //"corbaloc:iiop:localhost:9996/localserver";
        public static SolidEdgeFramework.Application objApp = null;
        static SolidEdgeFramework.SolidEdgeTCE objSEEC = null;

        public static SolidEdge.RevisionManager.Interop.Application objReviseApp = null;
        //static RevisionManager.Application objReviseApp1 = null;

        public static void SetCredentials(String username,string pass, string grp)
        {
            userName = username;
            password = pass;
            group = grp;
            Role = "";
        }

        public static void set_SE_Object (SolidEdgeFramework.Application app)
        {
            objApp = app;
            objSEEC = objApp.SolidEdgeTCE;
            log.write(logType.INFO, "objSEEC Acquired");
        }
        public static void LoginToTeamcenter()
        {
            //Get Active session of Solid Edge 
            log.write(logType.INFO,"Initating SEEC");
            //objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);

            log.write(logType.INFO, "Starting new solidedge session and logging in to download files");
            SolidEdgeFramework.Application objApp = null;
            objApp = (SolidEdgeFramework.Application)Activator.
                CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
            //if (objApp == null)
            //throw new Exception("Could not start new solid edge  session");
            if (objApp == null)
            {
                log.write(logType.INFO, "Could not start new solid edge  session");
                return;
            }

            // Teamcenter Mode
            try
            {
                objSEEC = objApp.SolidEdgeTCE;
               log.write(logType.INFO, "objSEEC Acquired");

                if (objSEEC == null)
                {
                    log.write(logType.INFO, "objSEEC is NULL");
                    return;
                }


                bool bTeamCenterMode = false;
                objSEEC.GetTeamCenterMode(out bTeamCenterMode);
                if (bTeamCenterMode == false)
                {
                    objApp.DisplayAlerts = false;
                    objSEEC.SetTeamCenterMode(true);
                }
                //string bStrCurrentUser = null;
                // objSEEC.GetCurrentUserName(out bStrCurrentUser);
                //if (bStrCurrentUser.Equals(""))
                //{
                log.write(logType.INFO, "Validating login...");
                objApp.DisplayAlerts = false;
                 objSEEC.ValidateLogin(userName, password, group, Role, URL);
                //}
                log.write(logType.INFO, "SEEC Login Successful to TC");
            }
            catch (Exception ex)
            {
                log.write(logType.ERROR,"SEEC login stopped. Using already loggined user details \n" + ex.ToString());
            }
        }

        public static SolidEdgeFramework.SolidEdgeTCE getSEECObject()
        {
            return objSEEC;
        }

        public static void KillSESession()
        {
            if (objApp == null) return;

            try
            {
                objApp.Quit();
            }
            catch (Exception ex)
            {
                log.write(logType.ERROR, "KillSESession Exception: " + ex.ToString());
            }
        }

       


        internal static string getRevisionID(string outputXLfileName)
        {
            String RevisionID = "";
            String itemID = "";
            if (objSEEC == null) return RevisionID;

            objSEEC.GetDocumentUID(outputXLfileName, out itemID, out RevisionID);

            return RevisionID;
        }

        internal static string GetPDMCachePath()
        {
            String cacheDir = "";
            if (objSEEC == null) return cacheDir;

            objSEEC.GetPDMCachePath(out cacheDir);
            return cacheDir;
        }

     

        //=================================================================================
        public static void InitializeSolidEdgeRevisionManagerSession()
        {
            try
            {
                log.write(logType.INFO, "SE_SESSION - Creating Revision Manager Instance ... ");
                objReviseApp = new SolidEdge.RevisionManager.Interop.Application();
                objReviseApp.DisplayAlerts = 0;
                log.write(logType.INFO, "SE_SESSION - Created...Revision Manager Instance .");

            }
            catch (Exception ex)
            {
                log.write(logType.ERROR,"SE_SESSION - " + ex.Message.ToString());
                log.write(logType.ERROR, ex.StackTrace.ToString());
                return;
            }
        }

        //==============================================================================
        public static void killRevisionManager()
        {
            objReviseApp.DisplayAlerts = 1;
            objReviseApp.Quit();
            objReviseApp = null;
        }

        public static void DeleteFilesFromCache( string cachePath)
        {
            log.write(logType.INFO, "SEEC Cache path is " + cachePath);
            log.write(logType.INFO, "Deleting files from cache");
            string[] fileNames = Directory.GetFiles(cachePath, "*.*", SearchOption.AllDirectories);
            foreach (string s in fileNames)
            {
                try
                {
                    if (s.EndsWith(".par") || s.EndsWith(".asm") || s.EndsWith(".pwd") || s.EndsWith(".psm") ||
                            s.EndsWith(".dft") || s.EndsWith(".xlsx") || s.EndsWith(".xls"))
                    {
                        System.Object[] arr = new System.Object[1];
                        arr[0] = s;
                        log.write(logType.INFO, "Deleting file from cache: " + s);
                        objSEEC.DeleteFilesFromCache(arr);
                    }
                }
                catch (Exception ex)
                {
                    log.write(logType.INFO, ex.ToString());
                    continue;
                }
            }
            log.write(logType.INFO, "Deleted files in cache");
        }
    }
}
