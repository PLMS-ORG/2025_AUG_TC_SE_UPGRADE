using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace _2DAutomaticUpdate
{
    class LTC_SEEC
    {
        public static SolidEdgeFramework.Application application = null;

        public static void DeleteFilesFromCache(SolidEdgeFramework.SolidEdgeTCE objSEEC, string cachePath, string logFilePath)
        {
            Utility.Log("SEEC Cache path is " + cachePath, logFilePath);
            Utility.Log("Deleting files from cache", logFilePath);
            string[] fileNames = Directory.GetFiles(cachePath, "*.*", SearchOption.AllDirectories);
            foreach (string s in fileNames)
            {
                try
                {
                    if (s.EndsWith(".par") || s.EndsWith(".asm") || s.EndsWith(".pwd") || s.EndsWith(".psm") ||
                            s.EndsWith(".dft"))
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


        public static SolidEdgeFramework.Application getSolidEdgeSession()
        {
            return application;
        }

        public static SolidEdgeFramework.Application Start()
        {
            Type type = null;
            try
            {
                // Get the type from the Solid Edge ProgID
                type = Type.GetTypeFromProgID("SolidEdge.Application");
                // Start Solid Edge
                application = (SolidEdgeFramework.Application)
                Activator.CreateInstance(type);
                // Make Solid Edge visible
                application.Visible = false;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return application;

        }

        public static void QuitSEEC(SolidEdgeFramework.Application objApp, String logFilePath)
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
             SolidEdgeFramework.SolidEdgeTCE objSEEC,
            SolidEdgeFramework.Application objApp, String logFilePath)
        {
            String fmsHome = System.Configuration.ConfigurationManager.AppSettings["FMS Home"];
            String userName = System.Configuration.ConfigurationManager.AppSettings["User Name"];
            String password = System.Configuration.ConfigurationManager.AppSettings["Password"];
            String group = System.Configuration.ConfigurationManager.AppSettings["Group"];
            String role = System.Configuration.ConfigurationManager.AppSettings["Role"];
            String URL = System.Configuration.ConfigurationManager.AppSettings["URL"];

            //objSEEC.ValidateLogin(userName, password, group, role, URL);

            System.Object vFileNames;
            int nFiles = 0;
            try
            {
                bool fileExists = false;
                objSEEC.DoesTeamCenterFileExists(itemIDRecieved, revIDReceived, out fileExists);
                if (fileExists == true)
                {
                    objApp.DisplayAlerts = false;
                    //objSEEC.ValidateLogin(userName, password, group, role, URL);
                    objSEEC.GetListOfFilesFromTeamcenterServer(itemIDRecieved, revIDReceived, out vFileNames, out nFiles);
                    System.Object[] objArray = (System.Object[])vFileNames;
                    bool foundasm = false;
                    string asyFileName = null;
                    bool founddft = false;
                    string drawingFileName = null;
                    string otherFileName = null;

                    foreach (System.Object o in objArray)
                    {
                        string filename = (string)(o);
                        Console.WriteLine("filename: " + filename);
                        if (filename.ToLower().Trim().EndsWith(".asm"))
                        {
                            foundasm = true;
                            asyFileName = filename;
                        }
                        else if (filename.ToLower().Trim().EndsWith(".dft"))
                        {
                            founddft = true;
                            drawingFileName = filename;
                        }
                        else if (filename.ToLower().Trim().EndsWith(".par") || filename.ToLower().Trim().EndsWith(".pwd") || filename.ToLower().Trim().EndsWith(".psm"))
                            otherFileName = filename;
                    }

                    if (founddft == true)
                    {
                        Console.WriteLine("Downloaded " + itemIDRecieved + "/" + revIDReceived + " with all levels including drawing");
                        System.Object[,] temp = new object[1, 1];
                        objApp.DisplayAlerts = false;
                        //objSEEC.ValidateLogin(userName, password, group, role, URL);
                        objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, drawingFileName,
                            SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, true,
                            1, temp);
                    }
                    else
                    {
                        if (foundasm == true)
                        {
                            Utility.Log("Downloaded " + itemIDRecieved + "/" + revIDReceived + " with all levels", logFilePath);
                            System.Object[,] temp = new object[1, 1];
                            objApp.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, asyFileName,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, true,
                                1, temp);
                        }
                        else if (otherFileName != null)
                        {
                            Utility.Log("Downloaded " + itemIDRecieved + "/" + revIDReceived + " for single level", logFilePath);
                            System.Object[,] temp = new object[1, 1];
                            objApp.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, otherFileName,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                1, temp);
                        }
                    }

                    int NoOfComponents = 0;
                    System.Object ListOfItemRevIds = null, ListOfFileSpecs = null;
                    //Console.WriteLine(itemIDRecieved);
                    //Console.WriteLine(revIDReceived);
                    //objSEEC = (SolidEdge.Framework.Interop.SolidEdgeTCE)objApp.SolidEdgeTCE;
                    //objApp.DisplayAlerts = false;
                    //objSEEC.SetTeamCenterMode(true);
                    //objSEEC.ValidateLogin(userName, password, group, role, URL);
                    objApp.DisplayAlerts = false;
                    //objSEEC.ValidateLogin(userName, password, group, role, URL);
                    objSEEC.GetBomStructure(itemIDRecieved, revIDReceived, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                        out ListOfItemRevIds, out ListOfFileSpecs);
                    Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                    itemAndRevIds.Add(itemIDRecieved, revIDReceived);
                    if (NoOfComponents > 0)
                    {
                        System.Object[,] abcd = (System.Object[,])ListOfItemRevIds;
                        for (int i = 0; i < abcd.GetUpperBound(0); i++)
                        {
                            if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                                itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                        }
                    }
                    vFileNames = null;
                    nFiles = 0;
                    objArray = null;
                    //Downloading all applicable drawings for current assembly
                    foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                    {
                        Utility.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath);
                        vFileNames = null;
                        nFiles = 0;
                        objApp.DisplayAlerts = false;
                        //objSEEC.ValidateLogin(userName, password, group, role, URL);
                        objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                        objArray = null;
                        objArray = (System.Object[])vFileNames;
                        foreach (System.Object o in objArray)
                        {
                            string filename = (string)(o);
                            if (filename.Contains(".dft"))
                            {
                                System.Object[,] temp = new object[1, 1];
                                objApp.DisplayAlerts = false;
                                //objSEEC.ValidateLogin(userName, password, group, role, URL);
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                    1, temp);
                                Utility.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
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
                               SolidEdgeFramework.Application application)
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

        public static void syncComponents(String newItemFilePath, String logFilePath)
        {
            List<string> uniqueOcc = new List<string>();

            SolidEdgeFramework.Application objApp = (SolidEdgeFramework.Application)SolidEdgeCommunity.SolidEdgeUtils.Connect();
            SolidEdgeFramework.SolidEdgeTCE objSEEC = (SolidEdgeFramework.SolidEdgeTCE)objApp.SolidEdgeTCE;

            SolidEdgeFramework.Documents objDocuments = objApp.Documents;
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
            objApp.DisplayAlerts = true;

        }


        public static void CheckInSEDocumentsToTeamcenter(string logFilePath, List<string> listOfFileNamesInSession)
        {
            String bstrCachePath = "";
            SolidEdgeFramework.SolidEdgeTCE objSEEC = (SolidEdgeFramework.SolidEdgeTCE)application.SolidEdgeTCE;;
            objSEEC.GetPDMCachePath(out bstrCachePath);

            String[] asmFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".asm", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();

            String[] parFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".par", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();

            String[] psmFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();

            String[] dftFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".dft", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();

            Utility.Log("Initial file count for check in", logFilePath);
            Utility.Log("asmFiles " + asmFiles.Length.ToString(), logFilePath);
            Utility.Log("parFiles " + parFiles.Length.ToString(), logFilePath);
            Utility.Log("psmFiles " + psmFiles.Length.ToString(), logFilePath);
            Utility.Log("dftFiles " + dftFiles.Length.ToString(), logFilePath);

            //Removing pars not applicable for this assy

            try
            {
                List<string> asmFilesTemp = new List<string>();
                foreach (String asmFile in asmFiles)
                {
                    if (listOfFileNamesInSession.Contains(Path.GetFileName(asmFile.ToLower().Trim())))
                    {
                        Utility.Log(asmFile + " is applicable for this assembly to be checked in", logFilePath);
                        asmFilesTemp.Add(asmFile);
                    }
                    else
                        Utility.Log(asmFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                asmFiles = asmFilesTemp.ToArray();
            }
            catch (Exception)
            {

            }


            //Removing pars not applicable for this assy
            try
            {
                List<string> parFilesTemp = new List<string>();
                foreach (String parFile in parFiles)
                {
                    if (listOfFileNamesInSession.Contains(Path.GetFileName(parFile.ToLower().Trim())))
                    {
                        Utility.Log(parFile + " is applicable for this assembly to be checked in", logFilePath);
                        parFilesTemp.Add(parFile);
                    }
                    else
                        Utility.Log(parFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                parFiles = parFilesTemp.ToArray();
            }
            catch (Exception)
            {

            }



            //Removing psms not applicable for this assy
            try
            {
                List<string> psmFilesTemp = new List<string>();
                foreach (String psmFile in psmFiles)
                {
                    if (listOfFileNamesInSession.Contains(Path.GetFileName(psmFile.ToLower().Trim())))
                    {
                        Utility.Log(psmFile + " is applicable for this assembly to be checked in", logFilePath);
                        psmFilesTemp.Add(psmFile);
                    }
                    else
                        Utility.Log(psmFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                psmFiles = psmFilesTemp.ToArray();
            }
            catch (Exception)
            {


            }



            //Removing drafts not applicable for this assy
            try
            {
                List<string> draftFilesTemp = new List<string>();
                foreach (String draftFile in dftFiles)
                {
                    if (listOfFileNamesInSession.Contains(Path.GetFileName(draftFile.ToLower().Trim())))
                    {
                        Utility.Log(draftFile + " is applicable for this assembly to be checked in", logFilePath);
                        draftFilesTemp.Add(draftFile);
                    }
                    else
                        Utility.Log(draftFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                dftFiles = draftFilesTemp.ToArray();
            }
            catch (Exception)
            {


            }

            Utility.Log("Final file count for check in", logFilePath);
            Utility.Log("asmFiles " + asmFiles.Length.ToString(), logFilePath);
            Utility.Log("parFiles " + parFiles.Length.ToString(), logFilePath);
            Utility.Log("psmFiles " + psmFiles.Length.ToString(), logFilePath);
            Utility.Log("dftFiles " + dftFiles.Length.ToString(), logFilePath);

            int i = 0;
            try
            {
                if (asmFiles.Length != 0)
                {
                    Array ppsaAssemFileList = Array.CreateInstance(typeof(object), asmFiles.Length);
                    i = 0;
                    foreach (String asemFile in asmFiles)
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(asemFile);
                        Utility.Log(asemFile + "...checkout Status...: " + ischeckedout, logFilePath);
                        ppsaAssemFileList.SetValue((object)asemFile, i);
                        i++;
                    }


                    application.DisplayAlerts = false;
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsaAssemFileList, false);
                    application.DisplayAlerts = true;
                }
                else
                    Utility.Log("No Assembly files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Exception while checking in assemblies " + ex.ToString(), logFilePath);

            }

            try
            {
                if (parFiles.Length != 0)
                {
                    Array ppsaPartsFileList = Array.CreateInstance(typeof(object), parFiles.Length);
                    i = 0;
                    foreach (String partFile in parFiles)
                    {
                        ppsaPartsFileList.SetValue((object)partFile, i);
                        i++;
                    }
                    application.DisplayAlerts = false;
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsaPartsFileList, false);
                    application.DisplayAlerts = true;
                }
                else
                    Utility.Log("No Part files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Exception while checking in parts " + ex.ToString(), logFilePath);
            }

            try
            {
                if (psmFiles.Length != 0)
                {
                    Array ppsaSheetMetalFileList = Array.CreateInstance(typeof(object), psmFiles.Length);
                    i = 0;
                    foreach (String psmFile in psmFiles)
                    {
                        ppsaSheetMetalFileList.SetValue((object)psmFile, i);
                        i++;
                    }
                    application.DisplayAlerts = false;
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsaSheetMetalFileList, false);
                    application.DisplayAlerts = true;
                }
                else
                    Utility.Log("No psm files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Exception while checking in sheet metals " + ex.ToString(), logFilePath);
            }

            try
            {
                if (dftFiles.Length != 0)
                {
                    Array ppsadftFileList = Array.CreateInstance(typeof(object), dftFiles.Length);
                    i = 0;
                    foreach (String dftFile in dftFiles)
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(dftFile);
                        Utility.Log(dftFile + "...checkout Status...: " + ischeckedout, logFilePath);
                        ppsadftFileList.SetValue((object)dftFile, i);
                        i++;
                    }


                    application.DisplayAlerts = false;
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsadftFileList, false);
                    application.DisplayAlerts = true;
                }
                else
                    Utility.Log("No draft files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Exception while checking in drafts " + ex.ToString(), logFilePath);
            }

            //objSEEC.DeleteFilesFromCache(ref ppsaAssemFileList);

        }

        public static SolidEdgeFramework.Application GetActiveObject()
        {

            SolidEdgeFramework.Application application1 = null;
            Type type = null;

            Process[] pname = Process.GetProcessesByName("Edge");
            if (pname.Length != 0)
            {

                application1 = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                Console.WriteLine("Got the Active Object");
                application = application1;
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
    }
}
