
using ExcelSyncTC.controller;
using ExcelSyncTC.model;
using ExcelSyncTC.opInterfaces;
using ExcelSyncTC.utils;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelSyncTC.TC;

namespace ExcelSyncTC
{
    public partial class ExcelSyncDialog : Form
    {
        public ExcelSyncDialog()
        {
            InitializeComponent();

            Trace.Listeners.Add(new ListBoxTraceListener(listBox1));
        }

        private void SyncTE_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;

            List<object> arguments = new List<object>();
            arguments.Add(xlApp);


            if (backgroundWorker1.IsBusy != true)
            {
                this.button1.Enabled = false;
                this.progressBar1.Visible = true;
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("SyncTE: Process is Already Running.");
                return;
            } 
        }



        private void SyncToSolidEdge(Microsoft.Office.Interop.Excel.Application xlApp,String topLineAssembly,String logFilePath)
        {
            
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            
            try
            {
                //Utlity.Log("Saving the Changes Done..", logFilePath, "INFO");
                xlApp.ActiveWorkbook.Save();

                Utlity.Log("Reading Data from Excel..", logFilePath,"INFO");
                ExcelData.readOccurenceVariablesFromTemplateExcelFast(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
                return;
            }

            //Utlity.Log("Connecting to Solid Edge..", logFilePath, "INFO");
            //SE_SESSION.InitializeSolidEdgeSession(logFilePath);

            try
            {
                Utlity.Log("Syncing Data to Solid Edge..", logFilePath, "INFO");
                SolidEdgeFramework.Application application = null;
                SolidEdgeDocument document = null;
                application = SE_SESSION.getSolidEdgeSession();
                if (application == null)
                {
                    MessageBox.Show("Solid Edge Application is NULL");
                    return;
                }
                document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

                if (document == null)
                {
                    MessageBox.Show("Solid Edge Document is NULL");
                    return;
                }
                topLineAssembly = document.FullName;
                Utlity.Log("topLineAssembly: " + topLineAssembly, logFilePath);
                SolidEdgeInterface.SolidEdgeSync(topLineAssembly, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
                return;
            }

           

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            Utlity.Log("Completed Syncing Data To Solid Edge..", logFilePath, "INFO");
            //MessageBox.Show("Sync to Solid Edge Completed");
            //xlApp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            //xlApp = null;
        }

        // REQUEST - 22 OCT
        private void SyncFeaturesToSolidEdge(Microsoft.Office.Interop.Excel.Application xlApp, String topLineAssembly, String logFilePath)
        {
            
            String xlFilePath = xlApp.ActiveWorkbook.FullName;
            Utlity.Log("SolidEdgeSetFeatureState: xlFilePath: " + xlFilePath, logFilePath);
            Utlity.Log("SolidEdgeSetFeatureState: readFeaturesFromTemplateExcel: " + xlFilePath, logFilePath);
            ExcelReadFeatures.readFeaturesFromTemplateExcel(xlApp, xlFilePath, logFilePath);
            Utlity.Log("SolidEdgeSetFeatureState: getFeatureLinesList: " + xlFilePath, logFilePath);
            List<FeatureLine> updatedFsList = ExcelReadFeatures.getFeatureLinesList();            
            try
            {
                Utlity.Log("SolidEdgeSetFeatureState: setFeatures: " + System.DateTime.Now.ToString(), logFilePath);
                SolidEdgeSetFeatureState SetFeature= new SolidEdgeSetFeatureState();
                Thread myThread = new Thread(() => SetFeature.SolidEdgeFeatureSyncFromExcel(topLineAssembly, updatedFsList, logFilePath));
                myThread.SetApartmentState(ApartmentState.STA);
                myThread.Start();
                myThread.Join();
            }
            catch (Exception ex)
            {                
                Utlity.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);                
                return;
            }
            Utlity.Log("SyncFeaturesToSolidEdge: Completed: " + System.DateTime.Now.ToString(), logFilePath);
        }

        /** Use Case - Remove Component - 29 OCT
         * User to Delete a Variable Part in "MASTER ASSEMBLY"
         * User to Click on SyncTE in Excel.
         * Procedure to Remove the Sheets Automatically.
         * User to Remove Component from Solid Edge Manually
         **/
        private void ReadComponentTabFromExcel(String logFilePath)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;
            try
            {
                Utlity.Log("ReadMasterAssemblySheet: ", logFilePath);
                MasterAssemblyReader.ReadMasterAssemblySheet(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("ReadMasterAssemblySheet: " + ex.Message, logFilePath);
                return;
            }

            try
            {
                Utlity.Log("RemoveSheet: ", logFilePath);
                ExcelRemoveComponent.RemoveSheet(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("RemoveSheet: " + ex.Message, logFilePath);
                return;
            }

            try
            {
                Utlity.Log("RemoveComponentsFromFeatureTab: ", logFilePath);
                ExcelRemoveComponent.RemoveComponentsInFeatureTAB(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("RemoveComponentsFromFeatureTab: " + ex.Message, logFilePath);
                return;
            }


        }

        public static void DownloadFilesToCache(SolidEdgeFramework.Application application, SolidEdgeFramework.SolidEdgeTCE objSEEC,
           SolidEdgeDocument Sedocument, string logFilePath)
        {
            
            //Finding drawings of current assembly
            try
            {
                objSEEC = application.SolidEdgeTCE;
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
                    for (int i = 0; i <= abcd.GetUpperBound(0); i++)
                    {
                        if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                            itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                    }
                }
                System.Object vFileNames = null;
                int nFiles = 0;
                System.Object[] objArray = null;
                foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                {
                    vFileNames = null;
                    nFiles = 0;
                    application.DisplayAlerts = false;
                    objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                    objArray = null;
                    objArray = (System.Object[])vFileNames;
                    foreach (System.Object o in objArray)
                    {
                        string filename = (string)(o);
                        if (filename.Contains(".dft"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                1, temp);
                            Utlity.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.Contains(".par"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded part of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);

                        }

                        if (filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                               SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded sheet metal of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded Assembly of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded weldment of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".par", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".dft", StringComparison.OrdinalIgnoreCase))
                        {

                            if (listOfFileNamesInSession.Contains(filename) == false)
                                listOfFileNamesInSession.Add(filename.ToLower().Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("An exception was caught while trying to fetch list of file names " + ex.ToString(), logFilePath);

            }
            if (listOfFileNamesInSession.Count != 0)
            {
                Utlity.Log("Printing listOfFileNamesInSession", logFilePath);
                foreach (string s in listOfFileNamesInSession)
                    Utlity.Log(s, logFilePath);
            }
            else
                Utlity.Log("No documents found in listOfFileNamesInSession", logFilePath);
        }

        public static List<string> listOfFileNamesInSession = new List<string>();
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (Microsoft.Office.Interop.Excel.Application)genericlist[0];

            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;


            if (xlApp.ActiveWorkbook == null)
            {
                MessageBox.Show("SyncTE: Cannot Access Excel Application");
                this.button1.Enabled = true;
                return;
            }
            String xlFilePath = xlApp.ActiveWorkbook.FullName;
            String LogStageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(xlFilePath) + "_" + "SyncTE" + ".txt");

            Utlity.Log("SyncTE 3D Started", logFilePath);
            Utlity.Log("Connecting to Solid Edge..", logFilePath, "INFO");
            SE_SESSION.InitializeSolidEdgeSession(logFilePath);
            SolidEdgeFramework.Application Seapplication = null;
            SolidEdgeFramework.SolidEdgeTCE objSEEC = null;
            SolidEdgeDocument Sedocument = null;
            Seapplication = SE_SESSION.getSolidEdgeSession();
            if (Seapplication == null)
            {
                MessageBox.Show("Solid Edge Application is NULL");
                return;
            }
            Sedocument = (SolidEdgeFramework.SolidEdgeDocument)Seapplication.ActiveDocument;

            if (Sedocument == null)
            {
                MessageBox.Show("Solid Edge Document is NULL");
                return;
            }
            String topLineAssembly = Sedocument.FullName;
            Utlity.Log("topLineAssembly: " + topLineAssembly, logFilePath);
            //String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");

            //Finding drawings of current assembly
            try
            {
                objSEEC = Seapplication.SolidEdgeTCE;
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
                    for (int i = 0; i <= abcd.GetUpperBound(0); i++)
                    {
                        if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                            itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                    }
                }
                System.Object vFileNames = null;
                int nFiles = 0;
                System.Object[] objArray = null;
                foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                {
                    vFileNames = null;
                    nFiles = 0;
                    application.DisplayAlerts = false;
                    objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                    objArray = null;
                    objArray = (System.Object[])vFileNames;
                    foreach (System.Object o in objArray)
                    {
                        string filename = (string)(o);
                        if (filename.Contains(".dft"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                1, temp);
                            Utlity.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.Contains(".par"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded part of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                            
                        }

                        if (filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                               SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded sheet metal of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                           //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded Assembly of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            //application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded weldment of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".par", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".dft", StringComparison.OrdinalIgnoreCase))
                        {
                            
                            if (listOfFileNamesInSession.Contains(filename) == false)
                                listOfFileNamesInSession.Add(filename.ToLower().Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("An exception was caught while trying to fetch list of file names " + ex.ToString(), logFilePath);

            }
            if (listOfFileNamesInSession.Count != 0)
            {
                Utlity.Log("Printing listOfFileNamesInSession", logFilePath);
                foreach (string s in listOfFileNamesInSession)
                    Utlity.Log(s, logFilePath);
            }
            else
                Utlity.Log("No documents found in listOfFileNamesInSession", logFilePath);

            
            try
            {
                Utlity.Log("Saving the Changes Done..", logFilePath, "INFO");
                xlApp.ActiveWorkbook.Save();

                Utlity.Log("Remove Variable Parts (If Opted By User)..", logFilePath, "INFO");
                ReadComponentTabFromExcel(logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("ReadComponentTabFromExcel: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;

            }

            try
            {
                Utlity.Log("Deleting Occurrences In SolidEdge", logFilePath, "INFO");
                 SolidEdgeOccurenceDelete_1 occDelete = new SolidEdgeOccurenceDelete_1();
                 occDelete.SolidEdgeOccurrenceDeleteFromExcelSTAT(topLineAssembly, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SolidEdgeOccurrenceDeleteFromExcelSTAT: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }


            // Upload the files back to Teamcenter - 04-April-2022
            try
            {
                Utlity.Log("Uploading files back to Teamcenter after occurence delete....", logFilePath);
                SEECAdaptor.LoginToTeamcenter();
                SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath, listOfFileNamesInSession);

            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
                return;
            }

           if (application == null)
            {
                Utlity.Log("application object is null..", logFilePath);
                e.Result = "NOK";
                return;
            }

            objSEEC = Seapplication.SolidEdgeTCE;
            if (objSEEC == null)
            {
                Utlity.Log("SEEC object is null..", logFilePath);
                return;

            }

            String cachePath = "";
            objSEEC.GetPDMCachePath(out cachePath);
            Utlity.Log("RE-DownloadFilesToCache.." + cachePath, logFilePath);
            DownloadFilesToCache(Seapplication, objSEEC, (SolidEdgeFramework.SolidEdgeDocument)Seapplication.ActiveDocument,logFilePath);

            try
            {
                SyncToSolidEdge(xlApp,topLineAssembly,logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SyncToSolidEdge: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }

            // Moved the Feature Sync As Requested By Simone - 13 Dec 2018
            try
            {
                Utlity.Log("Syncing Features to Solid Edge", logFilePath, "INFO");
                SyncFeaturesToSolidEdge(xlApp, topLineAssembly, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SyncFeaturesToSolidEdge: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }

            Sedocument = (SolidEdgeFramework.SolidEdgeDocument)Seapplication.ActiveDocument;

            if (Sedocument == null)
            {
                MessageBox.Show("Solid Edge Document is NULL");
                return;
            }
            re_checkout_items_in_cache(Seapplication.SolidEdgeTCE, Sedocument, application, logFilePath);

            // Upload the files back to Teamcenter - 17 August 2019
            try
            {
                Utlity.Log("Uploading files back to Teamcenter....", logFilePath, "INFO");
                //TcAdaptor.login("dcproxy", "dcproxy", "Engineering", "Designer", logFilePath);
                //TcAdaptor.TcAdaptor_Init();
                SE_SESSION.InitializeSolidEdgeSession(logFilePath);
                SEECAdaptor.LoginToTeamcenter();
                SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath, ExcelSyncDialog.listOfFileNamesInSession);
                //TcAdaptor.uploadExcelToTC(xlFilePath, logFilePath);
                //TcAdaptor.logout(logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
                return;
            }

            //Utlity.Log("Completed SyncTE..", logFilePath, "INFO");
           String startString = "SyncTE 3D Started";
           bool parseFlag = utils.Utlity.parseLog(logFilePath, startString);
            if (parseFlag == false)
                e.Result = "NOK";
            else
                e.Result = "OK";
        }

        // Re checked out the items in Cache, In case they are cancelled out.
        private void re_checkout_items_in_cache( SolidEdgeFramework.SolidEdgeTCE objSEEC, SolidEdgeDocument Sedocument, Microsoft.Office.Interop.Excel.Application application, String logFilePath)
        {
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
                String bStrCachePath = "";
                objSEEC.GetPDMCachePath(out bStrCachePath);
                Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                itemAndRevIds.Add(documentNumberValue, revisionValue);
                if (NoOfComponents > 0)
                {
                    System.Object[,] abcd = (System.Object[,])ListOfItemRevIds;
                    for (int i = 0; i <= abcd.GetUpperBound(0); i++)
                    {
                        if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                            itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                    }
                }
                System.Object vFileNames = null;
                int nFiles = 0;
                System.Object[] objArray = null;
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
                        String fileNameFull = Path.Combine(bStrCachePath, filename);
                        if (filename.Contains(".dft"))
                        {
                            System.Object[,] temp = new object[1, 1];

                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileNameFull);
                            Utlity.Log(fileNameFull + "...checkout Status...: " + ischeckedout, logFilePath);
                            if (ischeckedout == 0)
                            {
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                    1, temp);
                                Utlity.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                            }
                        }

                        if (filename.Contains(".par"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            

                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileNameFull);
                            Utlity.Log(fileNameFull + "...checkout Status...: " + ischeckedout, logFilePath);
                            if (ischeckedout == 0)
                            {
                                application.DisplayAlerts = false;
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                                Utlity.Log("Downloaded part of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                            }

                           
                        }

                        if (filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];

                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileNameFull);
                            Utlity.Log(fileNameFull + "...checkout Status...: " + ischeckedout, logFilePath);
                            if (ischeckedout == 0)
                            {
                                application.DisplayAlerts = false;
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                   SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utlity.Log("Downloaded sheet metal of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                            }
                        }

                        if (filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileNameFull);
                            Utlity.Log(fileNameFull + "...checkout Status...: " + ischeckedout, logFilePath);
                            if (ischeckedout == 0)
                            {
                                application.DisplayAlerts = false;
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utlity.Log("Downloaded Assembly of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                            }
                        }

                        if (filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];

                            int ischeckedout = objSEEC.IsTeamCenterFileCheckedOut(fileNameFull);
                            Utlity.Log(fileNameFull + "...checkout Status...: " + ischeckedout, logFilePath);
                            if (ischeckedout == 0)
                            {
                                application.DisplayAlerts = false;
                                objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                    SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                    1, temp);
                                Utlity.Log("Downloaded weldment of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                            }
                        }

                        if (filename.EndsWith(".par", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) || filename.EndsWith(".dft", StringComparison.OrdinalIgnoreCase))
                        {

                            if (listOfFileNamesInSession.Contains(filename) == false)
                                listOfFileNamesInSession.Add(filename.ToLower().Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("An exception was caught while trying to fetch list of file names " + ex.ToString(), logFilePath);

            }
        }

        private void ModifyXL(Microsoft.Office.Interop.Excel.Application xlApp, String xlFilePath,String logFilePath)
        {
            Utlity.Log("Updating Template Excel with Latest Values From Solid Edge..", logFilePath, "INFO");
            try
            {
                ExcelInterface.SaveDeltaToXL(xlApp, xlFilePath, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SaveDeltaToXL: " + ex.Message, logFilePath);
            }
            return;
        }

        [STAThread]
        private void ConnectToSolidEdge(String logFilePath)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }

            SE_SESSION.setSolidEdgeSession(application);
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

            if (document == null)
            {
                MessageBox.Show("document is NULL");
                return;
            }
            //MessageBox.Show(document.FullName);

            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            String AssemblyDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String xlFileName = System.IO.Path.ChangeExtension(assemblyFileName, "xlsx");
            

            String stageDir = Utlity.CreateLogDirectory();          

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            try
            {
                Utlity.Log("--readVariablesForEachOccurence-- ", logFilePath);
                Utlity.Log("Reading Back Data From Solid Edge..", logFilePath, "INFO");
                SolidEdgeData1.readVariablesForEachOccurence(assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("readVariablesForEachOccurence: " + ex.Message, logFilePath);                
                return;
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result.Equals("OK") == true)
            {                
                MessageBox.Show("SyncTE completed");
                this.DialogResult = DialogResult.OK;
                this.Dispose();
                Utlity.ModSheetsInSession.Clear();
                return;
            }
            else
            {                
                MessageBox.Show("SyncTE Failed, Check Logs");
                this.DialogResult = DialogResult.Cancel;
                this.Dispose();
                Utlity.ModSheetsInSession.Clear();
                return;
            }

        }
    }
}
