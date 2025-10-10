using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using _3DAutomaticUpdate.utils;
using SolidEdge.Framework.Interop;
using _3DAutomaticUpdate.opInterfaces;
using _3DAutomaticUpdate.TC;
using _3DAutomaticUpdate.controller;

namespace _3DAutomaticUpdate
{
    // 04-08-2022: Fix for the issue in sync3D logic when it is combined with EXCLUDE part/sub assembly option.
    class Program
    {
        public static string fmsHome = null;
        public static string userName = null;
        public static string password = null;
        public static string group = null;
        public static string role = null;
        public static string URL = null;
        public static String itemIDRecieved = null;
        public static String revIDReceived = null;
        public static String taskFolderPath = null;
        [STAThread]
        static void Main(string[] args)
        {
            taskFolderPath = args[0];
            itemIDRecieved = args[1];
            revIDReceived = args[2];

            String logFilePath = System.IO.Path.Combine(taskFolderPath, "3DAutomaticUpdate" + "_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm", CultureInfo.InvariantCulture) + ".txt");


            Utility.Log("---Inputs----", logFilePath);
            Utility.Log("----------------------------", logFilePath);
            Utility.Log("TaskFolderPath: " + taskFolderPath, logFilePath);
            Utility.Log("itemID: " + itemIDRecieved, logFilePath);
            Utility.Log("RevID: " + revIDReceived, logFilePath);
            Utility.Log("----------------------------", logFilePath);

            SolidEdge.Framework.Interop.Application objApp = LTC_SEEC.GetActiveObject();
            //SolidEdge.Framework.Interop.Application objApp = LTC_SEEC.Start1();
            objApp.Visible = false;

            if (objApp == null)
            {
                Utility.Log("no application object created..", logFilePath);
                return;
            }


            //SolidEdge.Framework.Interop.SolidEdgeTCE objSEEC = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;
            SolidEdge.Framework.Interop.SolidEdgeTCE objSEEC = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;
            objApp.DisplayAlerts = false;
            objSEEC.SetTeamCenterMode(true);
            fmsHome = System.Configuration.ConfigurationManager.AppSettings["FMS Home"];
            userName = System.Configuration.ConfigurationManager.AppSettings["User Name"];
            password = System.Configuration.ConfigurationManager.AppSettings["Password"];
            group = System.Configuration.ConfigurationManager.AppSettings["Group"];
            role = System.Configuration.ConfigurationManager.AppSettings["Role"];
            URL = System.Configuration.ConfigurationManager.AppSettings["URL"];
            /*
                        fmsHome = @"D:\PLM\TC12_4_RC_PLMSServer\tccs";
                        userName = "plms";
                        password = "plms";
                        group = "Engineering";
                        role = "ChangeSpecialist1";
                        URL = "http://87.106.171.154/tc";
            */
            Utility.Log("fmsHome.." + fmsHome, logFilePath);
            Utility.Log("userName.." + userName, logFilePath);
            Utility.Log("group.." + group, logFilePath);
            Utility.Log("role.." + role, logFilePath);
            Utility.Log("URL.." + URL, logFilePath);

            objSEEC.ValidateLogin(userName, password, group, role, URL);
            Utility.Log("SEEC Log in Successful..", logFilePath);
            string cachePath = null;

            Utility.Log("SEEC Getting cache path..", logFilePath);
            objSEEC.GetPDMCachePath(out cachePath);

            Utility.Log("SEEC cache path.." + cachePath, logFilePath);
            LTC_SEEC.DeleteFilesFromCache(objSEEC, cachePath, logFilePath);

            Utility.Log("DownloadFilesIntoCache.." + cachePath, logFilePath);
            LTC_SEEC.DownloadFilesIntoCache(itemIDRecieved, revIDReceived, objSEEC, objApp, logFilePath);

            String assemblyFilePath = Path.Combine(cachePath, itemIDRecieved + ".asm");

            String xlFilePath = Path.Combine(cachePath, itemIDRecieved + ".xlsx");
            Utility.Log("syncComponents.." + assemblyFilePath, logFilePath);
            Utility.Log("xlFilePath:" + xlFilePath, logFilePath);
            try
            {

               
                syncExcel(logFilePath, assemblyFilePath, xlFilePath);

            }
            catch (Exception ex)
            {
                Utility.Log("Exception.." + ex.Message, logFilePath );
                Utility.QuitSEEC(objApp, logFilePath);
                return;
            }

            Utility.QuitSEEC(objApp, logFilePath);
        }
        public static List<string> listOfFileNamesInSession = new List<string>();
        public static void syncExcel(String logFilePath,String assemblyFilePath,String xlFilePath)
        {

            Microsoft.Office.Interop.Excel.Application sourcexlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sourcexlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks sourcexlworkbooks = null;
            string cachePath = Path.GetDirectoryName(assemblyFilePath);

            FileInfo sourcefile = new FileInfo(xlFilePath);
            File.SetAttributes(xlFilePath, FileAttributes.Normal);
            if (sourcefile.Exists == true)
            {
                sourcexlworkbooks = sourcexlApp.Workbooks;

                try
                {
                    sourcexlWorkbook = sourcexlworkbooks.Open(xlFilePath);

                    if (sourcexlApp.ActiveWorkbook == null)
                    {
                        Utility.Log("SyncTE: Cannot Access Excel Application", logFilePath);
                        return;
                    }

                    xlFilePath = sourcexlApp.ActiveWorkbook.FullName;
                    Utility.Log("Connecting to Solid Edge..", logFilePath);
                    SE_SESSION.InitializeSolidEdgeSession(logFilePath);

                    SolidEdgeFramework.Application Seapplication = null;
                    SolidEdgeDocument Sedocument = null;
                    Seapplication = SE_SESSION.getSolidEdgeSession();
                    if (Seapplication == null)
                    {
                        Utility.Log("Solid Edge Application is NULL", logFilePath);
                        return;
                    }
                    //  open document
                   

                    var documents = Seapplication.Documents;

                    Sedocument = (SolidEdgeDocument)documents.Open(assemblyFilePath);
                   
                   

                    if (Sedocument == null)
                    {
                        Utility.Log("Solid Edge Document is NULL", logFilePath);
                        return;
                    }


                    String topLineAssembly = Sedocument.FullName;
                    Utility.Log("topLineAssembly: " + topLineAssembly, logFilePath);

                    SolidEdgeFramework.SolidEdgeTCE objSEEC = Seapplication.SolidEdgeTCE;
                    //Finding drawings of current assembly
                    try
                    {
                        
                        SolidEdgeFramework.PropertySets propertySets = (SolidEdgeFramework.PropertySets)Sedocument.Properties;
                        SolidEdgeFramework.Properties projectInformation = (SolidEdgeFramework.Properties)propertySets.Item(5);
                        SolidEdgeFramework.Property revision = (SolidEdgeFramework.Property)projectInformation.Item(2);
                        string revisionValue = revision.get_Value().ToString();
                        SolidEdgeFramework.Property documentNumber = (SolidEdgeFramework.Property)projectInformation.Item(1);
                        string documentNumberValue = documentNumber.get_Value().ToString();
                        int NoOfComponents = 0;
                        System.Object ListOfItemRevIds = null, ListOfFileSpecs = null;
                        objSEEC.GetBomStructure(documentNumberValue, revisionValue, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                                out ListOfItemRevIds, out ListOfFileSpecs);
                        
                        Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                        itemAndRevIds.Add(documentNumberValue, revisionValue);
                        if (NoOfComponents > 0)
                        {
                            System.Object[,] abcd = (System.Object[,])ListOfItemRevIds;
                            for (int i = 0; i < abcd.GetUpperBound(0); i++)
                            {
                                if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                                    itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                            }
                        }
                       
                        if (listOfFileNamesInSession.Count != 0)
                        {
                            Utility.Log("Printing listOfFileNamesInSession", logFilePath);
                            foreach (string s in listOfFileNamesInSession)
                            {
                                Utility.Log(s, logFilePath);                                

                            }
                        }
                        else
                            Utility.Log("No documents found in listOfFileNamesInSession", logFilePath);

                      
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("An exception was caught while trying to fetch list of file names " + ex.ToString(), logFilePath);

                    }
                   


                    try
                    {
                        Utility.Log("Saving the Changes Done..", logFilePath);
                        sourcexlApp.DisplayAlerts = false;
                        sourcexlApp.ActiveWorkbook.Save();
                        sourcexlApp.DisplayAlerts = false;
                        Utility.Log("Remove Variable Parts (If Opted By User)..", logFilePath);
                        ExcelSync.ReadComponentTabFromExcel(sourcexlApp, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("ReadComponentTabFromExcel: " + ex.Message, logFilePath);
                        return;
                    }


                    try
                    {
                        Utility.Log("Deleting Occurrences In SolidEdge", logFilePath);
                        SolidEdgeOccurenceDelete_1 occDelete = new SolidEdgeOccurenceDelete_1();
                        occDelete.SolidEdgeOccurrenceDeleteFromExcelSTAT(topLineAssembly, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("SolidEdgeOccurrenceDeleteFromExcelSTAT: " + ex.Message, logFilePath);
                        return;
                    }

                    // Upload the files back to Teamcenter - 04-April-2022
                    try
                    {
                        Utility.Log("Uploading files back to Teamcenter after occurence delete....", logFilePath);                        
                        SEECAdaptor.LoginToTeamcenter();                       
                        SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath, Program.listOfFileNamesInSession);

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("Exception: " + ex.Message, logFilePath);
                        Utility.Log("Exception: " + ex.StackTrace, logFilePath);
                        return;
                    }

                    SolidEdge.Framework.Interop.Application objApp = LTC_SEEC.GetActiveObject();
                    SolidEdge.Framework.Interop.SolidEdgeTCE objSec = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;
                    if (objSec == null)
                    {
                        Utility.Log("SEEC object is null..", logFilePath);
                        return;

                    }                    

                    Utility.Log("RE-DownloadFilesIntoCache.." + cachePath, logFilePath);
                    LTC_SEEC.DownloadFilesIntoCache(itemIDRecieved, revIDReceived, objSec, objApp, logFilePath);


                    try
                    {
                        ExcelSync.SyncToSolidEdge(sourcexlApp, topLineAssembly, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("SyncToSolidEdge: " + ex.Message, logFilePath);
                        return;
                    }


                    // Moved the Feature Sync As Requested By Simone - 13 Dec 2018
                    try
                    {
                        Utility.Log("Syncing Features to Solid Edge", logFilePath);
                        ExcelSync.SyncFeaturesToSolidEdge(sourcexlApp, topLineAssembly, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("SyncFeaturesToSolidEdge: " + ex.Message, logFilePath);
                        return;
                    }
                   

                    sourcexlApp.Visible = false;
                    sourcexlApp.UserControl = false;
                    sourcexlWorkbook.Close(true);

                

                    Marshal.ReleaseComObject(sourcexlWorkbook);
                    sourcexlWorkbook = null;

                    Marshal.ReleaseComObject(sourcexlworkbooks);
                    sourcexlworkbooks = null;

                    sourcexlApp.DisplayAlerts = false;
                    sourcexlApp.Quit();

                    Marshal.ReleaseComObject(sourcexlApp);
                    sourcexlApp = null;

                    // Upload the files back to Teamcenter - 17 August 2019
                    try
                    {
                        Utility.Log("Uploading files back to Teamcenter....", logFilePath);                        
                        SEECAdaptor.LoginToTeamcenter();                       
                        SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath, Program.listOfFileNamesInSession);
                       
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("Exception: " + ex.Message, logFilePath);
                        Utility.Log("Exception: " + ex.StackTrace, logFilePath);
                        return;
                    }


                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Console.WriteLine("file does not Exist: ");
            }
        }
    }
}
