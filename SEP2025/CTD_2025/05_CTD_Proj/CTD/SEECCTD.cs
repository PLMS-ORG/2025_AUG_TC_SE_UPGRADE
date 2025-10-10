using DemoAddInTC.utils;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Log;
using DemoAddInTC.se;

namespace CTD
{
    internal class SEECCTD
    {
        public static (string, string, string,Dictionary<String, String>) perform_ctd(string bstrItemID, string bstrItemRevID,
            string userName, string password, string group, string role, string URL)
        {
            try
            {
                log.write(logType.INFO, "Starting new solidedge session and logging in to download files");
                SolidEdgeFramework.Application objApp = null;
                objApp = (SolidEdgeFramework.Application)Activator.
                    CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                if (objApp == null)
                    throw new Exception("Could not start new solid edge  session");

                SolidEdgeFramework.SolidEdgeTCE objSEEC = null;
                objSEEC = (SolidEdgeFramework.SolidEdgeTCE)objApp.SolidEdgeTCE;
                if (objSEEC == null)
                    throw new Exception("Could not get seec instance from solid edge session");
                else
                    log.write(logType.INFO, "solidedge started...");


                objApp.DisplayAlerts = false;
                objSEEC.SetTeamCenterMode(true);

                log.write(logType.INFO, "Login Details:");
                log.write(logType.INFO, "userName = " + userName);
                log.write(logType.INFO, "password = " + password);
                log.write(logType.INFO, "group = " + group);
                log.write(logType.INFO, "role = " + role);
                log.write(logType.INFO, "URL = " + URL);
                log.write(logType.INFO, "Validating Login:");

                objApp.DisplayAlerts = false;
                objSEEC.ValidateLogin(userName, password, group, role, URL);
                log.write(logType.INFO, "logged  in to TC ..");

                // clean up the cache before downloading the current assembly
                String cachePath1 = "";
                objSEEC.GetPDMCachePath(out cachePath1);
                //SEECAdaptor.DeleteFilesFromCache(cachePath1);

                //Getting BOM structure
                log.write(logType.INFO, "Getting bom structure of selected item " + bstrItemID);

                int NoOfComponents = 0;
                System.Object ListOfItemRevIds, ListOfFileSpecs;
                objSEEC.GetBomStructure(bstrItemID, bstrItemRevID,
                                        SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(),
                                        true, out NoOfComponents,
                                        out ListOfItemRevIds, out ListOfFileSpecs);

                Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                itemAndRevIds.Add(bstrItemID, bstrItemRevID);
                log.write(logType.INFO, "Number of components are  " + NoOfComponents);
                System.Object[,] listOfItemRevIds1 = (System.Object[,])ListOfItemRevIds;
                log.write(logType.INFO, "Number of ListOfItemRevIds   " + listOfItemRevIds1.Length);

                if (NoOfComponents > 0)
                {
                    log.write(logType.INFO, "list Of ItemRevIds of components:");

                    System.Object[,] listOfItemRevIds = (System.Object[,])ListOfItemRevIds;
                    for (int i = 0; i < listOfItemRevIds.GetUpperBound(0); i++)
                    {
                        if (itemAndRevIds.Keys.Contains(listOfItemRevIds[i, 0].ToString()) == false)
                        {
                            itemAndRevIds.Add(listOfItemRevIds[i, 0].ToString(), listOfItemRevIds[i, 1].ToString());
                            log.write(logType.INFO, listOfItemRevIds[i, 0].ToString() + "/" + listOfItemRevIds[i, 1].ToString());
                        }
                    }
                }

                //Getting filename of assembly items
                log.write(logType.INFO, "Getting filename of assembly items...");

                Dictionary<string, string> itemAndFileName = new Dictionary<string, string>();
                System.Object vFileNames = null;
                int nFiles = 0;
                System.Object[] objArray = null;
                log.write(logType.INFO, "filename of assembly items which contains (.par,.asm,.pwd,.psm\").");
                foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                {
                    vFileNames = null;
                    nFiles = 0;
                    objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                    objArray = null;
                    objArray = (System.Object[])vFileNames;

                    foreach (System.Object o in objArray)
                    {
                        string filename = (string)(o);
                        if (filename.Contains(".par") || filename.Contains(".asm") || filename.Contains(".pwd") || filename.Contains(".psm"))
                        {
                            if (itemAndFileName.Keys.Contains(pair.Key) == false)
                            {
                                itemAndFileName.Add(pair.Key, filename);
                                log.write(logType.INFO, pair.Key + ":" + filename);
                            }
                        }
                    }
                }

                //Starting structure editor
                log.write(logType.INFO, "Starting structure editor...");

                SolidEdge.StructureEditor.Interop.Application StructureEditorApplication = null;
                SolidEdge.StructureEditor.Interop.SEECStructureEditor SEECStructure = null;
                SolidEdge.StructureEditor.Interop.ISEECStructureEditorATP atp = null;
                SolidEdge.StructureEditor.Interop.ISEECStructureEditor istr = null;

                StructureEditorApplication = (SolidEdge.StructureEditor.Interop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("StructureEditor.Application"));
                if (StructureEditorApplication == null)
                    throw new Exception("Could not start structure editor. Make sure the application is installed");
                StructureEditorApplication.SetDisplayAlerts(false);
                StructureEditorApplication.Visible = 0;
                SEECStructure = StructureEditorApplication.SEECStructureEditor;
                if (SEECStructure == null)
                    throw new Exception("Could not get instance of structure editor. Make sure the application is installed");
                SEECStructure.Close();
                atp = StructureEditorApplication.SEECStructureEditorATP;
                if (atp == null)
                    throw new Exception("Could not get instance of structure editor atp. Make sure the application is installed");
                istr = StructureEditorApplication.SEECStructureEditor;
                if (istr == null)
                    throw new Exception("Could not get instance of ISEEC structure editor atp. Make sure the application is installed");

                //Logging In to Teamcenter from structure editor
                log.write(logType.INFO, "Logging In to Teamcenter from structure editor");

                int commandResult = -1;
                commandResult = SEECStructure.ValidateLogin(userName, password, group, role, URL);
                if (commandResult != 0)
                    throw new Exception("Could not log in to teamcenter from structure editor");
                else
                    log.write(logType.INFO, "Logging In to Teamcenter Successfully");

                commandResult = SEECStructure.ClearCache();
                if (commandResult != 0)
                    throw new Exception("Could not clear structure editor cache");

                //Opening item to clone in structure editor
                log.write(logType.INFO, "Opening item to clone in structure editor");

                String bstrFileName = "";
                String bstrRevisionRule = "Latest Working";
                String bstrFolderName = "";
                itemAndFileName.TryGetValue(bstrItemID, out bstrFileName);
                commandResult = SEECStructure.Open(bstrItemID, bstrItemRevID, bstrFileName, bstrRevisionRule, bstrFolderName);
                if (commandResult != 0)
                    throw new Exception("Cannot open " + bstrItemID + " in structure editor");

                //Settig action to save as all
                log.write(logType.INFO, "Settig action to save as all...");
                commandResult = SEECStructure.SetSaveAsAll();
                if (commandResult != 0)
                    throw new Exception("Cannot assign action to save as in structure editor");

                //Getting new item id for the top most assembly
                log.write(logType.INFO, "Getting new item id for the top most assembly...");
                string newItemID = null, newItemRevID = null;
                //if (modeOfWorking.Equals("debug"))
                //    objSEEC.AssignItemID("AI4_Item", out newItemID, out newItemRevID);
                //else
                //objSEEC.AssignItemID("AI4_Item", out newItemID, out newItemRevID);
                objSEEC.AssignItemID("Ltc4_Item", out newItemID, out newItemRevID);
                if (newItemID == null || newItemRevID == null)
                    throw new Exception("Could not fetch new item id and rev id from teamcenter");
                log.write(logType.INFO, "newItemRevID = " + newItemID + "/" + newItemRevID);

                //Setting item id for top most assembly
                log.write(logType.INFO, "Setting item id for top most assembly..." + bstrItemRevID + ":Filename =" + bstrFileName);
                commandResult = SEECStructure.SetDataIntoSingleCell(bstrItemID, bstrItemRevID, bstrFileName, "Item ID", newItemID);
                if (commandResult != 0)
                    throw new Exception("Error in Setting item id for top most assembly");
                commandResult = SEECStructure.SetDataIntoSingleCell(bstrItemID, bstrItemRevID, bstrFileName, "Revision", newItemRevID);
                if (commandResult != 0)
                    throw new Exception("Error in Setting item revision for top most assembly");
                log.write(logType.INFO, "Setting item id & revision  for top most assembly successful...");


                //Setting new item id to all items
                log.write(logType.INFO, "Setting new item id to all items...");
                commandResult = SEECStructure.AssignAll();
                if (commandResult != 0)
                    throw new Exception("Cannot assign new item ids in structure editor");
                log.write(logType.INFO, "Setting new item id to all items successful...");

                //Invoking run teamcenter validations
                log.write(logType.INFO, "Invoking run teamcenter validations...");
                System.Object pvarListOfErrors, pvarListOfWarnings;
                commandResult = istr.RunTeamcenterValidation(out pvarListOfErrors, out pvarListOfWarnings);
                if (pvarListOfErrors != null)
                    throw new Exception("Run Teamcenter validations failed");
                log.write(logType.INFO, " run teamcenter validations successful...");

                //Performing actions and getting TAL log file
                log.write(logType.INFO, "Performing actions and getting TAL log file...");
                string bstrTALLogFileName = null;
                atp.GetStructureEditorTALLogFileName(out bstrTALLogFileName);
                log.write(logType.INFO, "Performing actions and getting TAL log file succssful...");
                //MessageBox.Show(bstrTALLogFileName);
                commandResult = SEECStructure.PerformActions();
                if (commandResult != 0)
                    throw new Exception("Error encountered when performing actions");
                log.write(logType.INFO, "getting TAL log file succssful...= " + bstrTALLogFileName);
                log.write(logType.INFO, "Performing actions succssful...");

                //Terminating structure editor   
                log.write(logType.INFO, "Terminating structure editor");
                SEECStructure.Close();
                StructureEditorApplication.Quit();
                if (commandResult != 0)
                    throw new Exception("Error encountered when performing actions");
                log.write(logType.INFO, "Terminated structure editor");


                log.write(logType.INFO, "Checking if new assembly " + newItemID + " is already in cache " + newItemID);
                string cachePath = null;
                objSEEC.GetPDMCachePath(out cachePath);
                log.write(logType.INFO, "cachePath " + cachePath);
                if (cachePath == null)
                    throw new Exception("Could not get cache Path");
                vFileNames = null;
                nFiles = 0;
                objSEEC.GetListOfFilesFromTeamcenterServer(newItemID, newItemRevID, out vFileNames, out nFiles);
                objArray = null;
                objArray = (System.Object[])vFileNames;
                bool topLevelFileExistsInCache = false;
                string newItemFilePath = null;

                foreach (System.Object o in objArray)
                {
                    string filename = (string)(o);
                    log.write(logType.INFO, "filename " + filename);
                    if (filename.Contains(".par") || filename.Contains(".asm") || filename.Contains(".pwd") || filename.Contains(".psm"))
                    {
                        System.Object[,] temp = new object[1, 1];
                        if (System.IO.File.Exists(System.IO.Path.Combine(cachePath, filename)))
                        {
                            topLevelFileExistsInCache = true;
                            newItemFilePath = System.IO.Path.Combine(cachePath, filename);
                            log.write(logType.INFO, "The new item's file path in cache is " + newItemFilePath);
                        }
                        else
                        {
                            topLevelFileExistsInCache = false;
                            newItemFilePath = System.IO.Path.Combine(cachePath, filename);
                            log.write(logType.INFO, "The new item's file path in cache must be " + newItemFilePath);
                        }
                    }
                }

                if (topLevelFileExistsInCache == false)
                {
                    log.write(logType.INFO, "Getting BOM structure of " + newItemID);
                    NoOfComponents = 0;
                    ListOfItemRevIds = null; ListOfFileSpecs = null;
                    objSEEC.GetBomStructure(newItemID, newItemRevID, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                        out ListOfItemRevIds, out ListOfFileSpecs);
                    log.write(logType.INFO, "Number of components are  " + NoOfComponents);
                    if (NoOfComponents < 1)
                        log.write(logType.INFO, "Number of components are  zero");

                    itemAndRevIds = new Dictionary<string, string>();
                    itemAndRevIds.Add(newItemID, newItemRevID);
                    if (NoOfComponents > 0)
                    {
                        System.Object[,] tempObj = (System.Object[,])ListOfItemRevIds;
                        for (int i = 0; i < tempObj.GetUpperBound(0); i++)
                        {
                            if (itemAndRevIds.Keys.Contains(tempObj[i, 0].ToString()) == false)
                            {
                                itemAndRevIds.Add(tempObj[i, 0].ToString(), tempObj[i, 1].ToString());
                                log.write(logType.INFO, tempObj[i, 0].ToString() + "/" + tempObj[i, 1].ToString());

                            }
                        }
                    }

                    if (itemAndRevIds.Count() > 0)
                    {
                        log.write(logType.INFO, "Downloading all parts of " + newItemID);
                        if (cachePath == null)
                            throw new Exception("Could not get cache path");
                        foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                        {
                            log.write(logType.INFO, "Trying to download " + pair.Key + "/" + pair.Value);

                            vFileNames = null;
                            nFiles = 0;
                            objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                            objArray = null;
                            objArray = (System.Object[])vFileNames;
                            foreach (System.Object o in objArray)
                            {
                                string filename = (string)(o);
                                if (filename.Contains(".par") || filename.Contains(".asm") || filename.Contains(".pwd") || filename.Contains(".psm") ||
                                    filename.Contains(".dft"))
                                {
                                    System.Object[,] temp = new object[1, 1];
                                    objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                        SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                        1, temp);
                                    log.write(logType.INFO, "Downloaded " + pair.Key + "/" + pair.Value);
                                }
                            }
                        }
                    }
                }

                log.write(logType.INFO, "post clone EXcel sanitizing....");


                String[] XlFiles = Directory.GetFiles(cachePath, "*", SearchOption.AllDirectories)
                                                .Select(path => Path.GetFullPath(path))
                                                .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                                                .ToArray();
                foreach (String xLFile in XlFiles)
                {
                    log.write(logType.INFO, "Deleting xLFile: , " + xLFile);
                    System.IO.File.Delete(xLFile);
                }

                //log.write(logType.INFO, "Terminating solid edge..");
                //objApp.Quit();
                //log.write(logType.INFO, "Terminated solid edge..");

                SEECAdaptor.set_SE_Object(objApp);

               

                return (newItemID, newItemRevID, cachePath,itemAndRevIds);

            }
            catch (Exception ex)
            {
                log.write(logType.INFO, "exception in perform_ctd.." + ex.StackTrace + ex.Message);
                return ("", "", "", null);
            }

        }

       
        
    
    }
}

