using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace _3DAutomaticUpdate
{
   

    class LTC_SEEC
    {


        public static void DeleteFilesFromCache(SolidEdge.Framework.Interop.SolidEdgeTCE objSEEC, string cachePath, string logFilePath)
        {
            Utility.Log("SEEC Cache path is " + cachePath, logFilePath);
            Utility.Log("Deleting files from cache", logFilePath);           
            string[] fileNames = Directory.GetFiles(cachePath, "*.*", SearchOption.AllDirectories);
            foreach (string s in fileNames)
            {
                try
                {
                    if (s.EndsWith(".par") || s.EndsWith(".asm") || s.EndsWith(".pwd") || s.EndsWith(".psm") ||
                            s.EndsWith(".dft")|| s.EndsWith(".xlsx")||s.EndsWith(".xls"))
                    {
                        System.Object[] arr = new System.Object[1];
                        arr[0] = s;
                        Utility.Log("Deleting file from cache: " + s, logFilePath);
                        objSEEC.DeleteFilesFromCache(arr);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    continue;
                }
            }
            Utility.Log("Deleted files in cache", logFilePath);
        }



        public static SolidEdgeFramework.Application Start()
        {

            SolidEdgeFramework.Application application = null;
            Type type = null;
            try
            {
                // Get the type from the Solid Edge ProgID
                type = Type.GetTypeFromProgID("SolidEdge.Application");
                // Start Solid Edge
                application = (SolidEdgeFramework.Application) Activator.CreateInstance(type);
                // Make Solid Edge visible
                application.Visible = false;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return application;

        }

        public static SolidEdge.Framework.Interop.Application GetActiveObject()
        {

            SolidEdge.Framework.Interop.Application application = null;
            Type type = null;

             Process[] pname = Process.GetProcessesByName("Edge");
             if (pname.Length != 0)
             {

                 application = (SolidEdge.Framework.Interop.Application)Marshal.GetActiveObject("SolidEdge.Application");
                 Console.WriteLine("Got the Active Object");
             }
             else
             {
                 //try
                 //{
                 //    // Get the type from the Solid Edge ProgID
                 //    type = Type.GetTypeFromProgID("SolidEdge.Application");
                 //    // Start Solid Edge
                 //    application = (SolidEdgeFramework.Application)Activator.CreateInstance(type);
                 //    // Make Solid Edge visible
                 //    application.Visible = false;
                 //}
                 //catch (System.Exception ex)
                 //{
                 //    Console.WriteLine(ex.Message);
                 //}
             }
            return application;

        }

public static SolidEdge.Framework.Interop.Application Start1()
{

   SolidEdge.Framework.Interop.Application application = null;
    Type type = null;
    try
    {
        // Get the type from the Solid Edge ProgID
        type = Type.GetTypeFromProgID("SolidEdge.Application");
        // Start Solid Edge
        application = (SolidEdge.Framework.Interop.Application)
        Activator.CreateInstance(type);
        // Make Solid Edge visible
        application.Visible = false;
        application.ShowStartupScreen = false;
    }
    catch (System.Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
    return application;

}

public static void QuitSEEC(SolidEdge.Framework.Interop.Application objApp, String logFilePath)
        {
            try
            {
                objApp.Quit();
            }
            catch (Exception ex)
            {
                Utility.Log("QuitSEEC Exception.." + ex.Message, logFilePath);
                Utility.Log("QuitSEEC Exception.." + ex.StackTrace, logFilePath);
            }


        }

        public static void DownloadFilesIntoCache(String itemIDRecieved, String revIDReceived,
             SolidEdge.Framework.Interop.SolidEdgeTCE objSEEC,
            SolidEdge.Framework.Interop.Application objApp, String logFilePath)
        {
                     String fmsHome = System.Configuration.ConfigurationManager.AppSettings["FMS Home"];
                        String userName = System.Configuration.ConfigurationManager.AppSettings["User Name"];
                        String password = System.Configuration.ConfigurationManager.AppSettings["Password"];
                        String group = System.Configuration.ConfigurationManager.AppSettings["Group"];
                        String role = System.Configuration.ConfigurationManager.AppSettings["Role"];
                        String URL = System.Configuration.ConfigurationManager.AppSettings["URL"];
           
            System.Object vFileNames;
            int nFiles = 0;
            try
            {
                bool fileExists = false;
                objSEEC.DoesTeamCenterFileExists(itemIDRecieved, revIDReceived, out fileExists);
                if (fileExists == true)
                {
                    objApp.DisplayAlerts = false;
                    objSEEC.ValidateLogin(userName, password, group, role, URL);
                    objSEEC.GetListOfFilesFromTeamcenterServer(itemIDRecieved, revIDReceived, out vFileNames, out nFiles);
                    System.Object[] objArray = (System.Object[])vFileNames;
                    bool foundasm = false;
                    string asyFileName = null;
                    bool founddft = false;
                    string drawingFileName = null;
                    string otherFileName = null;
                    bool foundxls = false;
                    string xlsFileName = null;

                    foreach (System.Object o in objArray)
                    {
                        string filename = (string)(o);
                        Console.WriteLine("filename: " + filename);
                        if (filename.ToLower().Trim().EndsWith(".asm"))
                        {
                            foundasm = true;
                            asyFileName = filename;
                            if (Program.listOfFileNamesInSession.Contains(filename) == false)
                                Program.listOfFileNamesInSession.Add(filename.ToLower().Trim());
                        }
                        else if (filename.ToLower().Trim().EndsWith(".dft"))
                        {
                            founddft = true;
                            drawingFileName = filename;
                            if (Program.listOfFileNamesInSession.Contains(filename) == false)
                                Program.listOfFileNamesInSession.Add(filename.ToLower().Trim());
                        }
                        else if (filename.ToLower().Trim().EndsWith(".xls")|| filename.ToLower().Trim().EndsWith(".xlsx"))
                        {
                            foundxls = true;
                            xlsFileName = filename;
                           
                        }
                        else if (filename.ToLower().Trim().EndsWith(".par") || filename.ToLower().Trim().EndsWith(".pwd") || filename.ToLower().Trim().EndsWith(".psm"))
                            otherFileName = filename;
                        if (Program.listOfFileNamesInSession.Contains(filename) == false)
                            Program.listOfFileNamesInSession.Add(filename.ToLower().Trim());
                    }
                    if (foundxls == true)
                    {
                        Console.WriteLine("Downloaded " + itemIDRecieved + "/" + revIDReceived + " with all levels");
                        System.Object[,] temp = new object[1, 1];
                        objApp.DisplayAlerts = false;

                        objSEEC.ValidateLogin(userName, password, group, role, URL);
                        objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, xlsFileName,
                            SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", true, true,
                            1, temp);
                    }
                    if (founddft == true)
                    {
                        Console.WriteLine("Downloaded " + itemIDRecieved + "/" + revIDReceived + " with all levels including drawing");
                        System.Object[,] temp = new object[1, 1];
                        objApp.DisplayAlerts = false;
                       
                        objSEEC.ValidateLogin(userName, password, group, role, URL);
                        objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, drawingFileName,
                            SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", true, true,
                            1, temp);
                    }
                    else
                    {
                        if (foundasm == true)
                        {
                            Utility.Log("Downloaded " + itemIDRecieved + "/" + revIDReceived + " with all levels", logFilePath);
                            System.Object[,] temp = new object[1, 1];
                            objApp.DisplayAlerts = false;
                            objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, asyFileName,
                                SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                        }
                        else if (otherFileName != null)
                        {
                            Utility.Log("Downloaded " + itemIDRecieved + "/" + revIDReceived + " for single level", logFilePath);
                            System.Object[,] temp = new object[1, 1];
                            objApp.DisplayAlerts = false;
                            objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, otherFileName,
                                SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                1, temp);
                        }
                    }

                    int NoOfComponents = 0;
                    System.Object ListOfItemRevIds = null, ListOfFileSpecs = null;
                    Console.WriteLine(itemIDRecieved);
                    Console.WriteLine(revIDReceived);
                    //objSEEC = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;
                    objApp.DisplayAlerts = false;
                    //objSEEC.SetTeamCenterMode(true);
                    //objSEEC.ValidateLogin(userName, password, group, role, URL);
                    //objApp.DisplayAlerts = false;
                    objSEEC.ValidateLogin(userName, password, group, role, URL);
                    objSEEC.GetBomStructure(itemIDRecieved, revIDReceived, SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                        out ListOfItemRevIds, out ListOfFileSpecs);
                    Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                    itemAndRevIds.Add(itemIDRecieved, revIDReceived);
                    if (NoOfComponents > 0)
                    {
                        System.Object[,] abcd = (System.Object[,])ListOfItemRevIds;
                        for (int i = 0; i <= abcd.GetUpperBound(0); i++)
                        {
                            Utility.Log("itemAndRevId : " + abcd[i, 0].ToString(), logFilePath);
                            if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                                itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                        }
                    }
                    vFileNames = null;
                    nFiles = 0;
                    objArray = null;
                    Utility.Log("NoOfComponents: " + NoOfComponents, logFilePath);
                    //Utility.Log("itemAndRevIds: " + itemAndRevIds.Keys.Count, logFilePath);

                    //Downloading all applicable drawings for current assembly
                    foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                    {
                        Utility.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath);
                        vFileNames = null;
                        nFiles = 0;
                        objApp.DisplayAlerts = false;
                        objSEEC.ValidateLogin(userName, password, group, role, URL);
                        objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);

                        objArray = null;
                        objArray = (System.Object[])vFileNames;
                        foreach (System.Object o in objArray)
                        {
                            string filename = (string)(o);
                            Utility.Log("FileName: " + filename, logFilePath);
                            if (Program.listOfFileNamesInSession.Contains(filename) == false)
                                Program.listOfFileNamesInSession.Add(filename.ToLower().Trim());
                            if (filename.Contains(".dft"))
                            {
                                System.Object[,] temp = new object[1, 1];
                                objApp.DisplayAlerts = false;
                                //objSEEC.ValidateLogin(userName, password, group, role, URL);
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utility.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                                int isTCFileCheckedOut = objSEEC.IsTeamCenterFileCheckedOut(filename);
                                Utility.Log("Checkout Status of " + pair.Key + "/" + pair.Value + "is " + isTCFileCheckedOut, logFilePath);
                            }

                            if (filename.Contains(".par"))
                            {
                                System.Object[,] temp = new object[1, 1];
                                objApp.DisplayAlerts = false;
                                //objSEEC.ValidateLogin(userName, password, group, role, URL);
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utility.Log("Downloaded part of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                                int isTCFileCheckedOut = objSEEC.IsTeamCenterFileCheckedOut(filename);
                                Utility.Log("Checkout Status of " + pair.Key + "/" + pair.Value + "is " + isTCFileCheckedOut, logFilePath);
                            }

                            if (filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                System.Object[,] temp = new object[1, 1];
                                objApp.DisplayAlerts = false;
                                //objSEEC.ValidateLogin(userName, password, group, role, URL);
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utility.Log("Downloaded sheet metal of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                                int isTCFileCheckedOut = objSEEC.IsTeamCenterFileCheckedOut(filename);
                                Utility.Log("Checkout Status of " + pair.Key + "/" + pair.Value + "is " + isTCFileCheckedOut, logFilePath);
                            }

                            if (filename.EndsWith(".asm",StringComparison.OrdinalIgnoreCase) == true)
                            {
                                System.Object[,] temp = new object[1, 1];
                                objApp.DisplayAlerts = false;
                                //objSEEC.ValidateLogin(userName, password, group, role, URL);
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utility.Log("Downloaded Assembly of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                                int isTCFileCheckedOut = objSEEC.IsTeamCenterFileCheckedOut(filename);
                                Utility.Log("Checkout Status of " + pair.Key + "/" + pair.Value + "is " + isTCFileCheckedOut, logFilePath);
                            }

                            if (filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                System.Object[,] temp = new object[1, 1];
                                objApp.DisplayAlerts = false;
                                //objSEEC.ValidateLogin(userName, password, group, role, URL);
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdge.Framework.Interop.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utility.Log("Downloaded sheetmetal of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                                int isTCFileCheckedOut = objSEEC.IsTeamCenterFileCheckedOut(filename);
                                Utility.Log("Checkout Status of " + pair.Key + "/" + pair.Value + "is " + isTCFileCheckedOut, logFilePath);
                            }
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                // if there is an exception in download
                try
                {
                    objApp.Quit();
                }
                catch (Exception)
                {

                }
                Utility.Log("LTC_SE2CACHE: EXCEPTION \n" + ex.ToString() + "\n" + ex.StackTrace.ToString(), logFilePath);
                return;
            }

        }


        public static void getOccurence(List<string> uniqueOcc,
                             SolidEdgeAssembly.AssemblyDocument asyDoc,
                               SolidEdge.Framework.Interop.Application application)
        {
            foreach (SolidEdgeAssembly.Occurrence o in asyDoc.Occurrences)
            {
                string name = o.OccurrenceFileName;
                if (uniqueOcc.Contains(name) == false)
                {
                    uniqueOcc.Add(name);
                    if (name.ToLower().Trim().EndsWith(".asm"))
                    {
                        SolidEdgeAssembly.AssemblyDocument asyDoc1 = (SolidEdgeAssembly.AssemblyDocument)o.OccurrenceDocument;
                        getOccurence(uniqueOcc, asyDoc1, application);
                    }
                }
            }
        }

        public static void syncComponents(SolidEdge.Framework.Interop.Application objApp,
            String newItemFilePath, String logFilePath)
        {
            List<string> uniqueOcc = new List<string>();

            //SolidEdgeFramework.Application objApp = (SolidEdgeFramework.Application)SolidEdgeCommunity.SolidEdgeUtils.Connect();
            SolidEdge.Framework.Interop.SolidEdgeTCE objSEEC = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;

            SolidEdge.Framework.Interop.Documents objDocuments = objApp.Documents;
            if (objDocuments == null)
                throw new Exception("Object documents could not be loaded");

            objApp.DisplayAlerts = false;

            SolidEdgeAssembly.AssemblyDocument asyDoc = null;
            if (File.Exists(newItemFilePath) == true)
            {
                Utility.Log("Opening the new item in current solidedge session to check in the properties", logFilePath);
                objDocuments.Open(newItemFilePath);
            }
            else
            {
                Utility.Log("File Missing in cache.." + newItemFilePath, logFilePath);
                throw new Exception("File Missing in cache.." + newItemFilePath);
            }

            try
            {
                Utility.Log("Trying to get occurences", logFilePath);
                asyDoc = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;

                if (asyDoc.ReadOnly == true)
                {
                    bool WriteAccess = false;
                    asyDoc.SeekWriteAccess(out WriteAccess);
                    if (WriteAccess == false)
                    {
                        Utility.Log("Could not get WriteAccess to--" + newItemFilePath, logFilePath);
                        return;
                    }
                }
                getOccurence(uniqueOcc, asyDoc, objApp);
                Utility.Log("Printing occurences", logFilePath);
                foreach (string s in uniqueOcc)
                {
                    Utility.Log(s, logFilePath);
                }
                Utility.Log("Converting to object", logFilePath);
                System.Object[] objA1 = new object[1];
                objA1[0] = ((System.Object)asyDoc.FullName);
                Utility.Log("Checking in " + asyDoc.FullName, logFilePath);
                objSEEC.CheckInDocumentsToTeamCenterServer(objA1, false);
                asyDoc.Close();

                foreach (string s in uniqueOcc)
                {
                    Utility.Log("Opening " + s, logFilePath);
                    SolidEdgeFramework.SolidEdgeDocument d = (SolidEdgeFramework.SolidEdgeDocument)objDocuments.Open(s);
                    Utility.Log("Opened " + s, logFilePath);
                    System.Object[] objA = new object[1];
                    objA[0] = ((System.Object)s);
                    objSEEC.CheckInDocumentsToTeamCenterServer(objA, false);
                    Utility.Log("Checked in " + s, logFilePath);
                    d.Close();
                    Utility.Log("Closed " + s, logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Exception caught when trying to get occurences" + ex.ToString(), logFilePath);
            }
            objApp.DisplayAlerts = false;

        }
    }
}
