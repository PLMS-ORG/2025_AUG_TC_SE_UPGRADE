using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelSyncTC.utils;
using SolidEdge.StructureEditor.Interop;

namespace ExcelSyncTC.TC
{
    class SEECAdaptor
    {
        static String userName = "infodba";
        static String password = "infodba";
        static String group = "DBA";
        static String Role = "DBA";
        static String URL = "http://tc11nx10:8080/tc";
        static SolidEdgeFramework.Application objApp = null;
        static SolidEdgeFramework.SolidEdgeTCE objSEEC = null;

        public static void LoginToTeamcenter()
        {
            //Get Active session of Solid Edge 
            Console.WriteLine("Initating Solid Edge Application");
            objApp = SE_SESSION.getSolidEdgeSession();
            //objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Connect(true, false);
            if (objApp == null)
            {

                return;
            }


            // Teamcenter Mode
            objSEEC = objApp.SolidEdgeTCE;

            bool bTeamCenterMode = false;
            objSEEC.GetTeamCenterMode(out bTeamCenterMode);
            if (bTeamCenterMode == false)
            {

                objSEEC.SetTeamCenterMode(true);
            }

            objSEEC.ValidateLogin(userName, password, group, Role, URL);
            Console.WriteLine("Login Successful to TC");
        }

        public static void uploadTemplateExcelToTC(string fileName)
        {
            //SolidEdgeFramework.Application objApplication = null;
            //SolidEdgeFramework.SolidEdgeTCE ObjSEEC = null;
            SolidEdgeFramework.SolidEdgeDocument ObjDoc = null;
            object[,] ListOfPropsForFileSaveAs = new object[4, 2];
            string oldFileName = null;
            string newFileName = null;
            //string App_Path = null;

            try
            {
                newFileName = "";

                ListOfPropsForFileSaveAs[0, 0] = "Item ID";
                ListOfPropsForFileSaveAs[0, 1] = "000055";
                ListOfPropsForFileSaveAs[1, 0] = "Revision";
                ListOfPropsForFileSaveAs[1, 1] = "A";
                ListOfPropsForFileSaveAs[2, 0] = "Item Name";
                ListOfPropsForFileSaveAs[2, 1] = "000055 - A";
                ListOfPropsForFileSaveAs[3, 0] = "Dataset Name";
                ListOfPropsForFileSaveAs[3, 1] = "000055/A";
                object objectListOfProps = (object)ListOfPropsForFileSaveAs;

                objSEEC.SetPDMProperties(fileName, ref objectListOfProps, out newFileName);
                //ObjDoc = (SolidEdgeFramework.SolidEdgeDocument)objApp.Documents.Open(oldFileName);
                //ObjDoc.SaveAs(newFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        // tried out -- not working
        public static bool uploadTemplateFileToTeamcenter(String fileName)
        {

            Array ppsaFileList = Array.CreateInstance(typeof(object), 1);
            String fileNameWoExtn = System.IO.Path.GetFileName(fileName);
            //object[] ppsaFileList = new string[] { fileName };
            ppsaFileList.SetValue((object)fileName, 0);
            if (objSEEC == null) return false;
            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileName);
            Console.WriteLine("IsTeamCenterFileCheckedOut...: " + ischeckedout);
            Console.WriteLine("CheckInDocumentsToTeamCenterServer...to Teamcenter: " + fileName);
            //if (ischeckedout != 0)
            {
                //String bstrItemType = "";
                String bstrItemID = System.IO.Path.GetFileNameWithoutExtension(fileName);
                String bstrItemRevID = "A";
                //objSEEC.AssignItemID(bstrItemType, out bstrItemID, out bstrItemRevID);
                //objApp.Visible = false;
                object[,] temp = new object[1, 1];


                //objSEEC.CheckOutDocumentsFromTeamCenterServer(bstrItemID, bstrItemRevID, false, fileName, DocumentDownloadLevel.SEECDownloadAllLevel);
                objSEEC.CheckInDocumentsToTeamCenterServer(ppsaFileList, false);

                //objApp.Visible = true;

                //String bstrDataSetFileName = "000056.xlsx";
                //String bstrRevisionRule = "";
                //String cachePath = "";
                //objSEEC.GetPDMCachePath(out cachePath);
                //object pvarList = null;
                //objSEEC.SaveAsToTeamCenter("000056", "A", bstrDataSetFileName, bstrRevisionRule, cachePath, out pvarList);
                if (objSEEC == null) return false;
                object pVarListofOutOfDateDocuments = null;
                objSEEC.GetOutOfDateDocuments(out pVarListofOutOfDateDocuments);

            }
            return true;
        }

        public static void CheckInSEDocumentsToTeamcenter(string logFilePath, List<string> listOfFileNamesInSession)
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

            Utlity.Log("Initial file count for check in", logFilePath);
            Utlity.Log("asmFiles " + asmFiles.Length.ToString(), logFilePath);
            Utlity.Log("parFiles " + parFiles.Length.ToString(), logFilePath);
            Utlity.Log("psmFiles " + psmFiles.Length.ToString(), logFilePath);
            Utlity.Log("dftFiles " + dftFiles.Length.ToString(), logFilePath);

            //Removing pars not applicable for this assy

            try
            {
                List<string> asmFilesTemp = new List<string>();
                foreach (String asmFile in asmFiles)
                {
                    if (listOfFileNamesInSession.Contains(Path.GetFileName(asmFile.ToLower().Trim())))
                    {
                        Utlity.Log(asmFile + " is applicable for this assembly to be checked in", logFilePath);
                        asmFilesTemp.Add(asmFile);
                    }
                    else
                        Utlity.Log(asmFile + " is not applicable for this assembly to be checked in", logFilePath);
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
                        Utlity.Log(parFile + " is applicable for this assembly to be checked in", logFilePath);
                        parFilesTemp.Add(parFile);
                    }
                    else
                        Utlity.Log(parFile + " is not applicable for this assembly to be checked in", logFilePath);
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
                        Utlity.Log(psmFile + " is applicable for this assembly to be checked in", logFilePath);
                        psmFilesTemp.Add(psmFile);
                    }
                    else
                        Utlity.Log(psmFile + " is not applicable for this assembly to be checked in", logFilePath);
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
                        Utlity.Log(draftFile + " is applicable for this assembly to be checked in", logFilePath);
                        draftFilesTemp.Add(draftFile);
                    }
                    else
                        Utlity.Log(draftFile + " is not applicable for this assembly to be checked in", logFilePath);
                }
                dftFiles = draftFilesTemp.ToArray();
            }
            catch (Exception)
            {
                
               
            }

            Utlity.Log("Final file count for check in", logFilePath);
            Utlity.Log("asmFiles " + asmFiles.Length.ToString(), logFilePath);
            Utlity.Log("parFiles " + parFiles.Length.ToString(), logFilePath);
            Utlity.Log("psmFiles " + psmFiles.Length.ToString(), logFilePath);
            Utlity.Log("dftFiles " + dftFiles.Length.ToString(), logFilePath);

            int i = 0;
           
            try
            {
                if (parFiles.Length != 0)
                {
                    Array ppsaPartsFileList = Array.CreateInstance(typeof(object), parFiles.Length);
                    i = 0;
                    foreach (String partFile in parFiles)
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(partFile);
                        Utlity.Log(partFile + "...checkout Status...: " + ischeckedout, logFilePath);
                        ppsaPartsFileList.SetValue((object)partFile, i);
                        i++;
                    }
                    objApp.DisplayAlerts = false;
                    DateTime currentDateTime = DateTime.Now;
                    string formattedDateTime = currentDateTime.ToString("HH:mm:ss tt");
                    //Utlity.Log("Parts: " + formattedDateTime,logFilePath);
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsaPartsFileList, false);
                    objApp.DisplayAlerts = true;
                }
                else
                    Utlity.Log("No Part files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception while checking in parts " + ex.ToString(), logFilePath);
            }

            try
            {
                if (psmFiles.Length != 0)
                {
                    Array ppsaSheetMetalFileList = Array.CreateInstance(typeof(object), psmFiles.Length);
                    i = 0;
                    foreach (String psmFile in psmFiles)
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(psmFile);
                        Utlity.Log(psmFile + "...checkout Status...: " + ischeckedout, logFilePath);
                        ppsaSheetMetalFileList.SetValue((object)psmFile, i);
                        i++;
                    }
                    objApp.DisplayAlerts = false;
                    DateTime currentDateTime = DateTime.Now;
                    string formattedDateTime = currentDateTime.ToString("HH:mm:ss tt");
                    //Utlity.Log("sheetMetal: " + formattedDateTime, logFilePath);
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsaSheetMetalFileList, false);
                    objApp.DisplayAlerts = true;
                }
                else
                    Utlity.Log("No psm files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception while checking in sheet metals " + ex.ToString(), logFilePath);
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
                        Utlity.Log(dftFile + "...checkout Status...: " + ischeckedout, logFilePath);
                        ppsadftFileList.SetValue((object)dftFile, i);
                        i++;
                    }


                    objApp.DisplayAlerts = false;
                    DateTime currentDateTime = DateTime.Now;
                    string formattedDateTime = currentDateTime.ToString("HH:mm:ss tt");
                    //Utlity.Log("sheetMetal: " + formattedDateTime, logFilePath);
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsadftFileList, false);
                    objApp.DisplayAlerts = true;
                }
                else
                    Utlity.Log("No draft files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception while checking in drafts " + ex.ToString(), logFilePath);
            }

            //objSEEC.DeleteFilesFromCache(ref ppsaAssemFileList);

            try
            {
                if (asmFiles.Length != 0)
                {
                    Array ppsaAssemFileList = Array.CreateInstance(typeof(object), asmFiles.Length);
                    i = 0;
                    foreach (String asemFile in asmFiles)
                    {
                        int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(asemFile);
                        Utlity.Log(asemFile + "...checkout Status...: " + ischeckedout, logFilePath);
                        ppsaAssemFileList.SetValue((object)asemFile, i);
                        i++;
                    }


                    objApp.DisplayAlerts = false;
                    DateTime currentDateTime = DateTime.Now;
                    string formattedDateTime = currentDateTime.ToString("HH:mm:ss tt");
                    //Utlity.Log("Assys: " + formattedDateTime, logFilePath);
                    objSEEC.CheckInDocumentsToTeamCenterServer(ppsaAssemFileList, false);
                    objApp.DisplayAlerts = true;
                }
                else
                    Utlity.Log("No Assembly files to check in", logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception while checking in assemblies " + ex.ToString(), logFilePath);

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

        public static void CloneDataInTeamcenter()
        {

            SolidEdge.StructureEditor.Interop.Application StructureEditorApplication;
            StructureEditorApplication = (SolidEdge.StructureEditor.Interop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("StructureEditor.Application"), true);
            if (StructureEditorApplication == null) return;


            SEECStructureEditor SEECStructure = StructureEditorApplication.SEECStructureEditor;
            if (SEECStructure == null) return;

            Console.WriteLine("Logging In to Teamcenter: ");
            int iret = SEECStructure.ValidateLogin("dcproxy", "dcproxy", "", "", "corbaloc:iiop:localhost:9996/localserver");
            Console.WriteLine("iRet: " + iret);
            String bstrItemID = "6099553";
            String bstrItemRevID = "04";
            String bstrFileName = "";
            String bstrRevisionRule = "Latest Working";
            String bstrFolderName = "";

            SEECStructure.Open(bstrItemID, bstrItemRevID, bstrFileName, bstrRevisionRule, bstrFolderName);
            SEECStructure.SetSaveAsAll();
            SEECStructure.AssignAll();

            //SEECStructure.Close();
            SEECStructure.PerformActions();

        }


        internal static string getRevisionID(string outputXLfileName)
        {
            String RevisionID = "";
            String itemID = "";
            if (objSEEC == null) return RevisionID;

            objSEEC.GetDocumentUID(outputXLfileName, out itemID, out RevisionID);

            return RevisionID;
        }
    }
}
