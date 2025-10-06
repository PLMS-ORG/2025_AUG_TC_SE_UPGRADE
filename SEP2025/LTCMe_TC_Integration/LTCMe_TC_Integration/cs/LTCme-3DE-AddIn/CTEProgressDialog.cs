using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.utils;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using SolidEdgeCommunity.Extensions;
using System.IO;
using DemoAddInTC.se;
using Creo_TC_Live_Integration.TcDataManagement;

namespace DemoAddInTC
{
    public partial class CTEProgressDialog : Form
    {
        public CTEProgressDialog()
        {
            InitializeComponent();

            String fileName = getTemplateFileName();

            //if (fileName == null || fileName.Equals("") == true)
            //{
            //    this.label2.Text = "No Template Available for Active Assembly In Solid Edge";
            //}
            //else
            //{
            //    bool eligible = checkFileEligibility(fileName);
            //    String AsmFileName = System.IO.Path.GetFileName(fileName);
            //    this.label2.Text = AsmFileName;

            //    if (eligible == false)
            //    {
            //        MessageBox.Show("Template Not Available for Active Assembly in Solid Edge");
            //        this.button1.Enabled = false;
            //    }
            //}

        }

        private bool checkFileEligibility(string FullName)
        {
            
            if (System.IO.File.Exists(FullName) == true)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        // Read Assembly
        private void button1_Click_1(object sender, EventArgs e)
        {
            this.progressBar1.Visible = true;
            this.button1.Enabled = false;
            ConnectToTemplateExcel();
        }

        private String getTemplateFileName()
        {
            String fileName = "";
            SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                this.progressBar1.Visible = false;
                MessageBox.Show("Application is NULL");
                return "";
            }

            SE_SESSION.setSolidEdgeSession(application);
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

            if (document == null)
            {
                this.progressBar1.Visible = false;
                MessageBox.Show("document is NULL");
                return "";
            }
            //MessageBox.Show(document.FullName);

            fileName = document.FullName;
            String XLStageDir = System.IO.Path.GetDirectoryName(fileName);
            String xlFile = System.IO.Path.Combine(XLStageDir, System.IO.Path.GetFileNameWithoutExtension(fileName) + ".xlsx");

            return xlFile;

        }

        private void ConnectToTemplateExcel()
        {

            SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                this.progressBar1.Visible = false;
                MessageBox.Show("Application is NULL");
                return;
            }

            SE_SESSION.setSolidEdgeSession(application);
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

            if (document == null)
            {
                this.progressBar1.Visible = false;
                MessageBox.Show("document is NULL");
                return;
            }
            //MessageBox.Show(document.FullName);

            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            String XLStageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "CTE" + ".txt");

            String xlFile = System.IO.Path.Combine(XLStageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            //if (System.IO.File.Exists(xlFile) == false)
            //{
                downloadExcelTemplateFromTeamcenter(assemblyFileName, logFilePath);
               // 26 - 11 - 2024 - MURALI - Removed SOA call & Included the SEEC Download API
                downloadExcelTemplateFromTeamcenter_1(application, document, logFilePath);
            //}
                
               

            if (System.IO.File.Exists(xlFile) == false)
            {
                this.progressBar1.Visible = false;                
                MessageBox.Show(xlFile + " is Missing.");
                Utlity.Log(xlFile + " is Missing.", logFilePath);
                return;
            }

            List<object> arguments = new List<object>();
            arguments.Add(xlFile);
            arguments.Add(logFilePath);
            arguments.Add(application);
            arguments.Add(checkBox1.Checked);

            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker Thread is Busy, CTE", logFilePath);
            }

        }


        // 26-11-2024 - MURALI - Removed SOA call
        private void downloadExcelTemplateFromTeamcenter_1(SolidEdgeFramework.Application application, SolidEdgeDocument document, string logFilePath)
        {
            //SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)currentSESession.ActiveDocument;

            if (document == null)
            {
                MessageBox.Show("document is NULL");
                return;
            }

            String assemblyFileName = document.FullName;
            if (File.Exists(assemblyFileName) == false)
            {
                MessageBox.Show("assemblyFileName is NULL");
                return;
            }

            SolidEdgeData1.setAssemblyFileName(assemblyFileName);

            //String stageDir = Utlity.CreateLogDirectory();
            //String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "DownloadXLfromTC" + ".txt");
            SolidEdgeFramework.SolidEdgeTCE objSEEC = application.SolidEdgeTCE;

            if (objSEEC == null)
            {
                MessageBox.Show("objSEEC is NULL");
                return;
            }

            SolidEdgeFramework.Application objApp = application;
            //SolidEdgeFramework.SolidEdgeTCE ObjSEEC = currentSETCEObject;

            if (objSEEC == null)
            {
                MessageBox.Show("objSEEC is NULL");
                return;
            }

            String cachePath = "";
            objSEEC.GetPDMCachePath(out cachePath);

            if (cachePath == null || cachePath == "" || cachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                cachePath = Path.GetDirectoryName(assemblyFileName);
            }

            object ListofFiles = null;
            string Location = cachePath;
            object[] SEFiletypeFilters = new object[1];

            object[] RelationFilters = new object[1];
            object[] ReferanceFilters = new object[1];
            object[] ExportFileTypeFilter = new object[1];
            int dwExpandSelectionOptions = 0;

            try
            {
                //Solid Edge File type extensions filter without dot(.)
                SEFiletypeFilters[0] = "asm";
                //Name referance filter
                ReferanceFilters[0] = "excel";
                //Traslated file type extension filter without dot(.)
                ExportFileTypeFilter[0] = "xlsx";
                //Relation filter
                RelationFilters[0] = "IMAN_specification";

                dwExpandSelectionOptions = (int)SolidEdgeConstants.ExpandSelectionOptions.IncludeComponentsFromAssemblies;

                Utlity.Log("ExtractTranslatedFilesOfActiveDocument: ", logFilePath);
                Utlity.Log("Location: " + Location, logFilePath);

                objSEEC.ExtractTranslatedFilesOfActiveDocument(Location, (uint)dwExpandSelectionOptions, SEFiletypeFilters, RelationFilters, ReferanceFilters, ExportFileTypeFilter, out ListofFiles);

                if (ListofFiles == null)
                {
                    Utlity.Log("ListofFiles is null", logFilePath);
                }
                if (ListofFiles is List<string>)
                {
                    object obj = new List<string>();
                    List<string> files = (List<string>)ListofFiles;
                    // Iterate and print each item
                    foreach (string item in files)
                    {
                        Utlity.Log(item, logFilePath);
                    }
                }
                else
                {
                    Utlity.Log("The object is not a List<string>.", logFilePath);
                }

                Utlity.Log("ExtractTranslatedFilesOfActiveDocument Ended ", logFilePath);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                Utlity.Log("ExtractTranslatedFilesOfActiveDocument Exception " + ex.Message, logFilePath);
                Utlity.Log("ExtractTranslatedFilesOfActiveDocument Exception " + ex.Source, logFilePath);
            }
            finally
            {

                //Reset the object definition
                //ObjSEEC = null;
                //objApp = null;
            }
        }

        public static void downloadExcelTemplateFromTeamcenter(string assemblyFileName,String logFilePath)
        {
            SEECAdaptor.LoginToTeamcenter(logFilePath);

            string bStrCurrentUser = null;
            SEECAdaptor.getSEECObject().GetCurrentUserName(out bStrCurrentUser);

            String password = bStrCurrentUser;

            //TcAdaptor Tc = new TcAdaptor();
            Utlity.Log("Downloading Excel Template from Teamcenter... " + System.DateTime.Now.ToString(), logFilePath);
            
            Utlity.Log("Logging into TC..for TVS..NOV 2024", logFilePath);
            Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath);
            Utlity.Log("ID=" + bStrCurrentUser, logFilePath);
            Utlity.Log("Group=DBA", logFilePath);
            Utlity.Log("Role=dba", logFilePath);
            TcAdaptor.login(bStrCurrentUser, password, "DBA", "dba", logFilePath);
            Utlity.Log("Initializing TC Services..", logFilePath);

            TcAdaptor.TcAdaptor_Init(logFilePath);
            //SEECAdaptor.LoginToTeamcenter(logFilePath);

            String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String stageDir = SEECAdaptor.GetPDMCachePath();
          
          
            DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, stageDir, true, logFilePath, true);
            TcAdaptor.logout(logFilePath);
            Utlity.Log("Downloading Excel Template from Teamcenter Completed... " + System.DateTime.Now.ToString(), logFilePath);
        }
        

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            String xlFile = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            bool Checked = (bool)genericlist[3];

            application = (SolidEdgeFramework.Application)genericlist[2];
            try
            {
                Utlity.Log("readOccurenceVariablesFromTemplateExcelFast: ", logFilePath);
                //ExcelData.readOccurenceVariablesFromTemplateExcel(xlFile, logFilePath);
                ExcelData.readOccurenceVariablesFromTemplateExcelFast(xlFile, logFilePath);
            }
            catch (Exception ex)
            {

                this.DialogResult = DialogResult.Cancel;
                Utlity.Log("readOccurenceVariablesFromTemplateExcelFast: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }
            try
            {
                Utlity.Log("readFeaturesFromTemplateExcel: ", logFilePath);
                ExcelReadFeatures.readFeaturesFromTemplateExcel(xlFile, logFilePath);
            }
            catch (Exception ex)
            {

                this.DialogResult = DialogResult.Cancel;
                Utlity.Log("readFeaturesFromTemplateExcel: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }

            try
            {
                Utlity.Log("readOccurencePathFromTemplateExcelFast: ", logFilePath);
                ExcelData.readOccurencePathFromTemplateExcelFast(xlFile, logFilePath);
            }
            catch (Exception ex)
            {

                this.DialogResult = DialogResult.Cancel;
                Utlity.Log("readOccurencePathFromTemplateExcelFast: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }
            Utlity.Log("readOccurencePathFromTemplateExcelFast completed", logFilePath);

            try
            {
                //if (SolidEdgeHighLighter.getOccurenceCount() == 0)
                //{
                    Utlity.Log("--readOccurences-- ", logFilePath);
                    SolidEdgeHighLighter.readOccurences(logFilePath);
                //}
            }
            catch (Exception ex)
            {

                this.DialogResult = DialogResult.Cancel;
                Utlity.Log("readOccurences: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }

            // To Run Only for New Part...
            if (Checked == true)
            {
                String assemblyFileName = Path.Combine(Path.GetDirectoryName(xlFile), Path.GetFileNameWithoutExtension(xlFile) + ".asm");
                try
                {
                    
                    if (File.Exists(assemblyFileName) == true)
                    {
                        Utlity.Log("--readVariablesForEachOccurence-- ", logFilePath);
                        SolidEdgeData1.readVariablesForEachOccurence(assemblyFileName, logFilePath);
                    }
                }
                catch (Exception ex)
                {
                    this.DialogResult = DialogResult.Cancel;
                    // this.Dispose();
                    Utlity.Log("readVariablesForEachOccurence: " + ex.Message, logFilePath);
                    e.Result = null;
                    return;
                }

                try
                {
                    if (File.Exists(assemblyFileName) == true)
                    {
                        Utlity.Log("--traverseAssembly SolidEdgeData1-- ", logFilePath);
                        SolidEdgeData1.traverseAssembly(assemblyFileName, logFilePath);
                    }
                }
                catch (Exception ex)
                {
                    this.DialogResult = DialogResult.Cancel;
                    //this.Dispose();
                    Utlity.Log("SolidEdgeData1, traverseAssembly: " + ex.Message, logFilePath);
                    e.Result = null;
                    return;
                }

                // Start a Thread here, Since SolidEdge Functionality Runs Only in STA MODE.
                try
                {
                    Utlity.Log("readFeatures: " + System.DateTime.Now.ToString(), logFilePath);
                    Thread myThread = new Thread(() => SolidEdgeReadFeature.readFeatures(logFilePath, "TVS"));
                    myThread.SetApartmentState(ApartmentState.STA);
                    myThread.Start();
                    myThread.Join();
                }
                catch (Exception ex)
                {
                    this.DialogResult = DialogResult.Cancel;

                    //this.Dispose();
                    Utlity.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);
                    e.Result = null;
                    return;
                }

                List<string> listFromSolidEdge = new List<string>(SolidEdgeData1.getVariablesDictionaryDetails().Keys);             
                List<String> listFromExcel = ExcelData.getOcurrenceList();
                if (listFromSolidEdge != null && listFromSolidEdge.Count > 0 && listFromExcel != null && listFromExcel.Count > 0)
                {
                    List<String> NewPartAddedList = listFromSolidEdge.Except(listFromExcel).ToList();

                    if (NewPartAddedList != null && NewPartAddedList.Count > 0)
                    {
                        foreach (String Part in NewPartAddedList)
                        {
                            List<Variable>variableDetails = null;
                            SolidEdgeData1.getVariablesDictionaryDetails().TryGetValue(Part, out variableDetails);
                            if (variableDetails != null && variableDetails.Count > 0)
                            {
                                ExcelData.getVariableDetails().AddRange(variableDetails);
                                if (ExcelData.getVariableDictionary().ContainsKey(Part) == false)
                                {
                                    ExcelData.getVariableDictionary().Add(Part, variableDetails);
                                }
                                ExcelData.getOcurrenceList().Add(Part);
                                List<BOMLine> ListFromSolidEdge = SolidEdgeData1.getBomLinesList();
                                if (ListFromSolidEdge != null && ListFromSolidEdge.Count > 0)
                                {
                                    List<BOMLine> ListFromExcelData = ExcelData.getBomLineList();
                                    if (ListFromExcelData != null && ListFromExcelData.Count > 0)
                                    {


                                        var BOMLineToBeAdded =  ListFromSolidEdge.Where(l2 => 
    !ListFromExcelData.Any(l1 => l1.AbsolutePath.Equals(l2.AbsolutePath,StringComparison.OrdinalIgnoreCase) && l1.FullName.Equals(l2.FullName,StringComparison.OrdinalIgnoreCase)));
                                        if (BOMLineToBeAdded != null)
                                        {
                                            List<BOMLine> ListOfBOMLinesToBeAdded = BOMLineToBeAdded.ToList();
                                            if (ListOfBOMLinesToBeAdded != null && ListOfBOMLinesToBeAdded.Count > 0)
                                            {
                                                ExcelData.getBomLineList().AddRange(ListOfBOMLinesToBeAdded);
                                            }
                                        }
                                    }
                                }
                            }

                            Dictionary<String, List<FeatureLine>> FeatureDictionary = SolidEdgeReadFeature.getFeatureDictionary();
                            if (FeatureDictionary != null && FeatureDictionary.Count > 0)
                            {
                                List<FeatureLine> ListOfFeatureLines = null;
                                FeatureDictionary.TryGetValue(Part, out ListOfFeatureLines);
                                if (ListOfFeatureLines != null && ListOfFeatureLines.Count > 0)
                                {
                                    ExcelReadFeatures.getFeatureLinesList().AddRange(ListOfFeatureLines);
                                    if (ExcelReadFeatures.getFeatureDictionary().ContainsKey(Part) == false)
                                    {
                                        ExcelReadFeatures.getFeatureDictionary().Add(Part, ListOfFeatureLines);
                                    }
                                }
                            }

                        }
                    }
                }


            }


            e.Result = genericlist;

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = new List<object>();
            try
            {
                genericlist = e.Result as List<object>;
            }
            catch (Exception)
            {
                this.DialogResult = DialogResult.OK;
                return;
            }
            
            String xlFile = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            application = (SolidEdgeFramework.Application)genericlist[2];

            this.DialogResult = DialogResult.OK;

            List<Variable> AllVariablesList = ExcelData.getVariableDetails();

            Utlity.Log("variableList Size: " + AllVariablesList.Count, logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);

           

        }

        private void CTCProgressDialog_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            About a = new About();
            a.ShowDialog();
        }

       
    }
}
