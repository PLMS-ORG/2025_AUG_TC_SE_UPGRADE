using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _2DAutomaticUpdate
{
    class Program
    {
        public static string fmsHome = null;
        public static string userName = null;
        public static string password = null;
        public static string group = null;
        public static string role = null;
        public static string URL = null;

        [STAThread]
        static void Main(string[] args)
        {
            String taskFolderPath = args[0];
            String itemIDRecieved = args[1];
            String revIDReceived = args[2];

            String logFilePath = logFilePath = System.IO.Path.Combine(taskFolderPath, "2DAutomaticUpdate" + "_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm", CultureInfo.InvariantCulture) + ".txt");


            Utility.Log("---Inputs----", logFilePath);
            Utility.Log("----------------------------", logFilePath);
            Utility.Log("TaskFolderPath: " + taskFolderPath, logFilePath);
            Utility.Log("itemID: " + itemIDRecieved, logFilePath);
            Utility.Log("RevID: " + revIDReceived, logFilePath);
            Utility.Log("----------------------------", logFilePath);

            //SolidEdgeFramework.Application objApp = LTC_SEEC.Start();
            SolidEdgeFramework.Application objApp = LTC_SEEC.GetActiveObject();
            if (objApp == null)
            {
                Utility.Log("no application object created..", logFilePath);
                return;
            }

            SolidEdgeFramework.SolidEdgeTCE objSEEC = (SolidEdgeFramework.SolidEdgeTCE)objApp.SolidEdgeTCE;
            //SolidEdge.Framework.Interop.SolidEdgeTCE objSEEC = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;
            objApp.DisplayAlerts = false;
            objSEEC.SetTeamCenterMode(true);
            fmsHome = System.Configuration.ConfigurationManager.AppSettings["FMS Home"];
            userName = System.Configuration.ConfigurationManager.AppSettings["User Name"];
            password = System.Configuration.ConfigurationManager.AppSettings["Password"];
            group = System.Configuration.ConfigurationManager.AppSettings["Group"];
            role = System.Configuration.ConfigurationManager.AppSettings["Role"];
            URL = System.Configuration.ConfigurationManager.AppSettings["URL"];

            objSEEC.ValidateLogin(userName, password, group, role, URL);
            Utility.Log("SEEC Log in Successful..", logFilePath);
            string cachePath = null;

            Utility.Log("SEEC Getting cache path..", logFilePath);
            objSEEC.GetPDMCachePath(out cachePath);

            Utility.Log("SEEC cache path.." + cachePath, logFilePath);
            LTC_SEEC.DeleteFilesFromCache(objSEEC, cachePath, logFilePath);

            Utility.Log("DownloadFilesIntoCache.." + cachePath, logFilePath);
            LTC_SEEC.DownloadFilesIntoCache(itemIDRecieved, revIDReceived, objSEEC, objApp, logFilePath);

            Utility.Log("SyncDwgs.." + cachePath, logFilePath);
            SyncDwg.SyncDwgs(itemIDRecieved, revIDReceived, logFilePath);

            Utility.QuitSEEC(objApp, logFilePath);
        }
    }
}
