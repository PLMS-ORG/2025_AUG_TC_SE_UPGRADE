using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DemoAddInTC.utils;
using SolidEdgeConstants;
using SolidEdge.StructureEditor.Interop;
using SolidEdgeFramework;

namespace DemoAddInTC.se
{
    public class SEECAdaptor
    {
        static String userName = loginFromSE.userName;
        static String password = loginFromSE.password;
        static String group = loginFromSE.group;
        static String Role = loginFromSE.role;
        public static String URL = loginFromSE.URL; //"corbaloc:iiop:localhost:9996/localserver";
        static SolidEdgeFramework.Application objApp = null;
        static SolidEdgeFramework.SolidEdgeTCE objSEEC = null;

        public static SolidEdgeFramework.Application GetApplication
        {  get { return objApp; } }

         
        public static void LoginToTeamcenter(String logFilePath)
        {
            //Get Active session of Solid Edge 
            Utility.Log("Initating SEEC", logFilePath);
            objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(); //SE_SESSION.getSolidEdgeSession();
            //objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Connect(true, false);
            if (objApp == null) return;
            Utility.Log("SEEC acquired", logFilePath);

            // Teamcenter Mode
            try
            {
                objSEEC = objApp.SolidEdgeTCE;
                Utility.Log("objSEEC Acquired", logFilePath);

                bool bTeamCenterMode = false;
                objSEEC.GetTeamCenterMode(out bTeamCenterMode);
                if (bTeamCenterMode == false)
                {
                    objSEEC.SetTeamCenterMode(true);
                }

                string bStrCurrentUser = null;
                objSEEC.GetCurrentUserName(out bStrCurrentUser);
                if (bStrCurrentUser.Equals(""))
                {
                    Utility.Log("SEEC ValidateLogin", logFilePath);
                    objSEEC.ValidateLogin(userName, password, group, Role, URL);
                }

                
                Utility.Log("SEEC Login Successful to TC", logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SEEC login stopped. Using already loggined user details \n" + ex.ToString(), logFilePath);
            }


        }

        public static SolidEdgeFramework.SolidEdgeTCE getSEECObject()
        {
            return objSEEC;
        }

        //public static void uploadTemplateExcelToTC(string fileName)
        //{
        //    //SolidEdgeFramework.Application objApplication = null;
        //    //SolidEdgeFramework.SolidEdgeTCE ObjSEEC = null;
        //    SolidEdgeFramework.SolidEdgeDocument ObjDoc = null;
        //    object[,] ListOfPropsForFileSaveAs = new object[4, 2];
        //    string oldFileName = null;
        //    string newFileName = null;
        //    //string App_Path = null;

        //    try
        //    {
        //        newFileName = "";

        //        ListOfPropsForFileSaveAs[0, 0] = "Item ID";
        //        ListOfPropsForFileSaveAs[0, 1] = "000055";
        //        ListOfPropsForFileSaveAs[1, 0] = "Revision";
        //        ListOfPropsForFileSaveAs[1, 1] = "A";
        //        ListOfPropsForFileSaveAs[2, 0] = "Item Name";
        //        ListOfPropsForFileSaveAs[2, 1] = "000055 - A";
        //        ListOfPropsForFileSaveAs[3, 0] = "Dataset Name";
        //        ListOfPropsForFileSaveAs[3, 1] = "000055/A";
        //        object objectListOfProps = (object)ListOfPropsForFileSaveAs;

        //        objSEEC.SetPDMProperties(fileName, ref objectListOfProps, out newFileName);
        //        //ObjDoc = (SolidEdgeFramework.SolidEdgeDocument)objApp.Documents.Open(oldFileName);
        //        //ObjDoc.SaveAs(newFileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.ToString());
        //    }

        //}

        // tried out -- not working
        //public static bool uploadTemplateFileToTeamcenter(String fileName)
        //{

        //    Array ppsaFileList = Array.CreateInstance(typeof(object), 1);
        //    String fileNameWoExtn = System.IO.Path.GetFileName(fileName);
        //    //object[] ppsaFileList = new string[] { fileName };
        //    ppsaFileList.SetValue((object)fileName, 0);
        //    if (objSEEC == null) return false;
        //    int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileName);
        //    Console.WriteLine("IsTeamCenterFileCheckedOut...: " + ischeckedout);
        //    Console.WriteLine("CheckInDocumentsToTeamCenterServer...to Teamcenter: " + fileName);
        //    //if (ischeckedout != 0)
        //    {
        //        //String bstrItemType = "";
        //        String bstrItemID = System.IO.Path.GetFileNameWithoutExtension(fileName);
        //        String bstrItemRevID = "A";
        //        //objSEEC.AssignItemID(bstrItemType, out bstrItemID, out bstrItemRevID);
        //        //objApp.Visible = false;
        //        object[,] temp = new object[1, 1];


        //        //objSEEC.CheckOutDocumentsFromTeamCenterServer(bstrItemID, bstrItemRevID, false, fileName, DocumentDownloadLevel.SEECDownloadAllLevel);
        //        objSEEC.CheckInDocumentsToTeamCenterServer(ppsaFileList, false);

        //        //objApp.Visible = true;

        //        //String bstrDataSetFileName = "000056.xlsx";
        //        //String bstrRevisionRule = "";
        //        //String cachePath = "";
        //        //objSEEC.GetPDMCachePath(out cachePath);
        //        //object pvarList = null;
        //        //objSEEC.SaveAsToTeamCenter("000056", "A", bstrDataSetFileName, bstrRevisionRule, cachePath, out pvarList);
        //        if (objSEEC == null) return false;
        //        object pVarListofOutOfDateDocuments = null;
        //        objSEEC.GetOutOfDateDocuments(out pVarListofOutOfDateDocuments);

        //    }
        //    return true;
        //}

        public static void CheckInSEDocumentsToTeamcenter(String logFilePath)
        {
            try
            {
                String bstrCachePath = "";
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

                Utlity.Log("collected all files for checkin from.. " + bstrCachePath, logFilePath);



                //Removing pars not applicable for this assy
                List<string> asmFilesTemp = new List<string>();
                foreach (String asmFile in asmFiles)
                {
                    if (SyncTEDialog.listOfFileNamesInSession.Contains(Path.GetFileName(asmFile.ToLower().Trim())))
                    {
                        Utility.Log(asmFile + " is applicable for this assembly to be checked in", logFilePath);
                        asmFilesTemp.Add(asmFile);
                    }
                    else
                        Utility.Log(asmFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                asmFiles = asmFilesTemp.ToArray();


                //Removing pars not applicable for this assy
                List<string> parFilesTemp = new List<string>();
                foreach (String parFile in parFiles)
                {
                    if (SyncTEDialog.listOfFileNamesInSession.Contains(Path.GetFileName(parFile.ToLower().Trim())))
                    {
                        Utility.Log(parFile + " is applicable for this assembly to be checked in", logFilePath);
                        parFilesTemp.Add(parFile);
                    }
                    else
                        Utility.Log(parFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                parFiles = parFilesTemp.ToArray();



                //Removing psms not applicable for this assy
                List<string> psmFilesTemp = new List<string>();
                foreach (String psmFile in psmFiles)
                {
                    if (SyncTEDialog.listOfFileNamesInSession.Contains(Path.GetFileName(psmFile.ToLower().Trim())))
                    {
                        Utility.Log(psmFile + " is applicable for this assembly to be checked in", logFilePath);
                        psmFilesTemp.Add(psmFile);
                    }
                    else
                        Utility.Log(psmFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                psmFiles = psmFilesTemp.ToArray();



                //Removing drafts not applicable for this assy
                List<string> draftFilesTemp = new List<string>();
                foreach (String draftFile in dftFiles)
                {
                    if (SyncTEDialog.listOfFileNamesInSession.Contains(Path.GetFileName(draftFile.ToLower().Trim())))
                    {
                        Utility.Log(draftFile + " is applicable for this assembly to be checked in", logFilePath);
                        draftFilesTemp.Add(draftFile);
                    }
                    else
                        Utility.Log(draftFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                dftFiles = draftFilesTemp.ToArray();



                Array ppsaAssemFileList = Array.CreateInstance(typeof(object), asmFiles.Length);
                int i = 0;
                foreach (String asemFile in asmFiles)
                {
                    try
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(asemFile);
                        Console.WriteLine(asemFile + "...checkout Status...: " + ischeckedout);
                        ppsaAssemFileList.SetValue((object)asemFile, i);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            Utility.Log("IsTeamCenterFileCheckedOut " + ex.ToString(), logFilePath);
                            System.Threading.Thread.Sleep(5000);
                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(asemFile);
                            Console.WriteLine(asemFile + "...checkout Status...: " + ischeckedout);
                            ppsaAssemFileList.SetValue((object)asemFile, i);
                        }
                        catch (Exception ex1)
                        {
                            Utility.Log("IsTeamCenterFileCheckedOut " + ex1.ToString(), logFilePath);
                            i++;
                            continue;
                        }
                    }
                    i++;
                }


                objApp.DisplayAlerts = false;
                objSEEC.CheckInDocumentsToTeamCenterServer(ppsaAssemFileList, false);
                objApp.DisplayAlerts = true;

                Array ppsaPartsFileList = Array.CreateInstance(typeof(object), parFiles.Length);
                i = 0;
                foreach (String partFile in parFiles)
                {
                    ppsaPartsFileList.SetValue((object)partFile, i);
                    i++;
                }
                objApp.DisplayAlerts = false;
                objSEEC.CheckInDocumentsToTeamCenterServer(ppsaPartsFileList, false);
                objApp.DisplayAlerts = true;

                Array ppsaSheetMetalFileList = Array.CreateInstance(typeof(object), psmFiles.Length);
                i = 0;
                foreach (String psmFile in psmFiles)
                {
                    ppsaSheetMetalFileList.SetValue((object)psmFile, i);
                    i++;
                }
                objApp.DisplayAlerts = false;
                objSEEC.CheckInDocumentsToTeamCenterServer(ppsaSheetMetalFileList, false);
               
                objApp.DisplayAlerts = true;

                Array ppsadftFileList = Array.CreateInstance(typeof(object), dftFiles.Length);
                i = 0;
                foreach (String dftFile in dftFiles)
                {
                    try
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(dftFile);
                        Console.WriteLine(dftFile + "...checkout Status...: " + ischeckedout);
                        ppsadftFileList.SetValue((object)dftFile, i);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            Utility.Log("IsTeamCenterFileCheckedOut " + ex.ToString(), logFilePath);
                            System.Threading.Thread.Sleep(5000);
                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(dftFile);
                            Console.WriteLine(dftFile + "...checkout Status...: " + ischeckedout);
                            ppsadftFileList.SetValue((object)dftFile, i);
                        }
                        catch (Exception ex1)
                        {
                            Utility.Log("IsTeamCenterFileCheckedOut " + ex1.ToString(), logFilePath);
                            i++;
                            continue;
                        }

                    }

                    i++;
                }


                objApp.DisplayAlerts = false;
                objSEEC.CheckInDocumentsToTeamCenterServer(ppsadftFileList, false);
                objApp.DisplayAlerts = true;

                //objSEEC.DeleteFilesFromCache(ref ppsaAssemFileList);
            }
            catch (Exception ex)
            {
                Utlity.Log("CheckInSEDocumentsToTeamcenter " + ex.ToString(), logFilePath);
                Utlity.Log("CheckInSEDocumentsToTeamcenter " + ex.Message, logFilePath);
                Utlity.Log("CheckInSEDocumentsToTeamcenter " + ex.StackTrace, logFilePath);
            }


        }

        public static void uploadExcelTemplateToTeamcenter()
        {
            // upload the excelx file to TC using TC SOA
            if (objSEEC == null) return;

            String bstrCachePath = "";
            objSEEC.GetPDMCachePath(out bstrCachePath);

            String[] XLFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                            .Select(path => Path.GetFullPath(path))
                                            .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                                            .ToArray();
            if (XLFiles == null || XLFiles.Length == 0) return;
            if (XLFiles.Length > 1) return;

            Array ppsaExcelFileList = Array.CreateInstance(typeof(object), 1);
            int i = 0;
            //foreach (String asemFile in asmFiles)
            {
                int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(XLFiles[0]);
                Console.WriteLine(XLFiles[0] + "...checkout Status...: " + ischeckedout);
                ppsaExcelFileList.SetValue((object)XLFiles[0], 0);
                i++;
            }
            String bStrUserName = "";
            objSEEC.GetCurrentUserName(out bStrUserName);
            objSEEC.ImportDocumentsToServer(1, ref ppsaExcelFileList, bStrUserName, password, group, Role, URL, true, false, false, false, true, false, false, false, null);
        }

        //public static void CloneDataInTeamcenter()
        //{

        //    SolidEdge.StructureEditor.Interop.Application StructureEditorApplication;
        //    StructureEditorApplication = (SolidEdge.StructureEditor.Interop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("StructureEditor.Application"), true);
        //    if (StructureEditorApplication == null) return;


        //    SEECStructureEditor SEECStructure = StructureEditorApplication.SEECStructureEditor;
        //    if (SEECStructure == null) return;

        //    Console.WriteLine("Logging In to Teamcenter: ");
        //    int iret = SEECStructure.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
        //    Console.WriteLine("iRet: " + iret);
        //    String bstrItemID = "6099553";
        //    String bstrItemRevID = "04";
        //    String bstrFileName = "";
        //    String bstrRevisionRule = "Latest Working";
        //    String bstrFolderName = "";

        //    SEECStructure.Open(bstrItemID, bstrItemRevID, bstrFileName, bstrRevisionRule, bstrFolderName);
        //    SEECStructure.SetSaveAsAll();
        //    SEECStructure.AssignAll();

        //    //SEECStructure.Close();
        //    SEECStructure.PerformActions();

        //}


        internal static string getRevisionID(string outputXLfileName)
        {
            String RevisionID = "";
            String itemID = "";
            if (objSEEC == null) return RevisionID;

            objSEEC.GetDocumentUID(outputXLfileName, out itemID, out RevisionID);

            return RevisionID;
        }

        internal static string getItemID(string outputXLfileName)
        {
            String RevisionID = "";
            String itemID = "";
            if (objSEEC == null) return RevisionID;

            objSEEC.GetDocumentUID(outputXLfileName, out itemID, out RevisionID);

            return itemID;
        }

        internal static string GetPDMCachePath(String logFilePath = "")
        {
            String cacheDir = "";
            if (objSEEC == null)
            {
                if (logFilePath != "" || logFilePath.Equals("") == false)
                {
                    Utility.Log("cacheDir is Empty since SEEC Object is NULL", logFilePath);
                }
                return cacheDir;
            }

            objSEEC.GetPDMCachePath(out cacheDir);

            if (logFilePath != "" || logFilePath.Equals("") == false)
            {
                Utility.Log("cacheDir from GetPDMCachePath=" + cacheDir, logFilePath);
            }

            return cacheDir;
        }

        public static void FindBasedOn(String logFilePath)
        {

        }

        internal static void collectDraftFilesFromCacheAndRegisterScaleFactor(string logFilePath)
        {
            string bstrCachePath = "";
            SEECAdaptor.objSEEC.GetPDMCachePath(out bstrCachePath);
            Utility.Log(string.Concat("bstrCachePath:", bstrCachePath), logFilePath);
            string[] array = (
                from path in Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                select Path.GetFullPath(path) into x
                where x.EndsWith(".dft", StringComparison.OrdinalIgnoreCase)
                select x).ToArray<string>();
            for (int i = 0; i < (int)array.Length; i++)
            {
                string dftFile = array[i];
                RegisterScaleFactor.registerScaleFactor(dftFile, logFilePath);
                Utlity.Log(string.Concat("RegisterScaleFactor completed For:: ", dftFile), logFilePath, null);
            }
        }

        public static void DownloadFilesIntoCache(string itemIDRecieved, string revIDReceived, string logFilePath)
        {
            object vFileNames = null;
            Array arrays;
            RevisionManager.RevisionRuleType revisionRuleType;
            int nFiles = 0;
            try
            {
                bool fileExists = false;
                Utility.Log(string.Concat("DownloadFilesIntoCache: ", itemIDRecieved), logFilePath);
                Utility.Log(string.Concat("DownloadFilesIntoCache: ", revIDReceived), logFilePath);
                SEECAdaptor.objSEEC.DoesTeamCenterFileExists(itemIDRecieved, revIDReceived, out fileExists);
                if (fileExists == true)
                {
                    Utility.Log(string.Concat("fileExists: ", fileExists.ToString()), logFilePath);
                    SEECAdaptor.objApp.DisplayAlerts = false;
                    SEECAdaptor.objSEEC.GetListOfFilesFromTeamcenterServer(itemIDRecieved, revIDReceived, out vFileNames, out nFiles);
                    object[] objArray = (object[])vFileNames;
                    bool foundasm = false;
                    string asyFileName = null;
                    bool founddft = false;
                    string drawingFileName = null;
                    string otherFileName = null;
                    object[] objArray1 = objArray;
                    for (int num = 0; num < (int)objArray1.Length; num++)
                    {
                        string filename = (string)objArray1[num];
                        Console.WriteLine(string.Concat("filename: ", filename));
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
                        else if ((filename.ToLower().Trim().EndsWith(".par") || filename.ToLower().Trim().EndsWith(".pwd") ? true : filename.ToLower().Trim().EndsWith(".psm")))
                        {
                            otherFileName = filename;
                        }
                    }
                    if (founddft)
                    {
                        Console.WriteLine(string.Concat(new string[] { "Downloaded ", itemIDRecieved, "/", revIDReceived, " with all levels including drawing" }));
                        object[,] temp = new object[1, 1];
                        SEECAdaptor.objApp.DisplayAlerts = false;
                        SolidEdgeTCE solidEdgeTCE = SEECAdaptor.objSEEC;
                        revisionRuleType = RevisionManager.RevisionRuleType.LatestRevision;
                        arrays = temp;
                        solidEdgeTCE.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, drawingFileName, revisionRuleType.ToString(), "", true, true, 1, ref arrays);
                    }
                    else if (foundasm)
                    {
                        Utility.Log(string.Concat(new string[] { "Downloaded ", itemIDRecieved, "/", revIDReceived, " with all levels" }), logFilePath);
                        object[,] temp = new object[1, 1];
                        SEECAdaptor.objApp.DisplayAlerts = false;
                        SolidEdgeTCE solidEdgeTCE1 = SEECAdaptor.objSEEC;
                        revisionRuleType = RevisionManager.RevisionRuleType.LatestRevision;
                        arrays = temp;
                        solidEdgeTCE1.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, asyFileName, revisionRuleType.ToString(), "", true, true, 1, ref arrays);
                    }
                    else if (otherFileName != null)
                    {
                        Utility.Log(string.Concat(new string[] { "Downloaded ", itemIDRecieved, "/", revIDReceived, " for single level" }), logFilePath);
                        object[,] temp = new object[1, 1];
                        SEECAdaptor.objApp.DisplayAlerts = false;
                        SolidEdgeTCE solidEdgeTCE2 = SEECAdaptor.objSEEC;
                        revisionRuleType = RevisionManager.RevisionRuleType.LatestRevision;
                        arrays = temp;
                        solidEdgeTCE2.DownladDocumentsFromServerWithOptions(itemIDRecieved, revIDReceived, otherFileName, revisionRuleType.ToString(), "", true, false, 1, ref arrays);
                    }
                    int NoOfComponents = 0;
                    object ListOfItemRevIds = null;
                    object ListOfFileSpecs = null;
                    SEECAdaptor.objApp.DisplayAlerts = false;
                    revisionRuleType = RevisionManager.RevisionRuleType.LatestRevision;
                    SEECAdaptor.objSEEC.GetBomStructure(itemIDRecieved, revIDReceived, revisionRuleType.ToString(), true, out NoOfComponents, out ListOfItemRevIds, out ListOfFileSpecs);
                    Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>()
                    {
                        { itemIDRecieved, revIDReceived }
                    };
                    if (NoOfComponents > 0)
                    {
                        object[,] abcd = (object[,])ListOfItemRevIds;
                        for (int i = 0; i < abcd.GetUpperBound(0); i++)
                        {
                            if (!itemAndRevIds.Keys.Contains<string>(abcd[i, 0].ToString()))
                            {
                                itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                            }
                        }
                    }
                    vFileNames = null;
                    nFiles = 0;
                    objArray = null;
                    foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                    {
                        Utility.Log(string.Concat("Trying to download ", pair.Key, "/", pair.Value), logFilePath);
                        vFileNames = null;
                        nFiles = 0;
                        SEECAdaptor.objApp.DisplayAlerts = false;
                        SEECAdaptor.objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                        objArray = null;
                        object[] objArray2 = (object[])vFileNames;
                        for (int k = 0; k < (int)objArray2.Length; k++)
                        {
                            string filename = (string)objArray2[k];
                            if (filename.Contains(".dft"))
                            {
                                object[,] temp = new object[1, 1];
                                SEECAdaptor.objApp.DisplayAlerts = false;
                                SolidEdgeTCE solidEdgeTCE3 = SEECAdaptor.objSEEC;
                                string key = pair.Key;
                                string value = pair.Value;
                                revisionRuleType = RevisionManager.RevisionRuleType.LatestRevision;
                                arrays = temp;
                                solidEdgeTCE3.DownladDocumentsFromServerWithOptions(key, value, filename, revisionRuleType.ToString(), "", true, false, 1, ref arrays);
                                Utility.Log(string.Concat(new string[] { "Downloaded drawing of ", pair.Key, "/", pair.Value, " to cache" }), logFilePath);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               
                Utility.Log(string.Concat("LTC_SE2CACHE: EXCEPTION \n", ex.ToString(), "\n", ex.StackTrace.ToString()), logFilePath);
            }
        }
    }
}
