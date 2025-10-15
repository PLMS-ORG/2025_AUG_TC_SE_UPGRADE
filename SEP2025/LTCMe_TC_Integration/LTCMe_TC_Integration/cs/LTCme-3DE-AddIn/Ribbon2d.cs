using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.ComponentModel;
using System.Threading;
using System.Windows.Threading;

using DemoAddInTC.controller;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.utils;
using DemoAddInTC.model;
using DemoAddInTC.services;
using DemoAddInTC.se;
using Creo_TC_Live_Integration.TcDataManagement;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;
using Teamcenter.Services.Strong.Core._2008_06.DataManagement;
using Teamcenter.Services.Strong.Query._2010_09.SavedQuery;
using DemoAddInTC.TC;

namespace DemoAddInTC
{

    class Ribbon2d : SolidEdgeCommunity.AddIn.Ribbon
    {
        const string _embeddedResourceName = "DemoAddInTC.Ribbon2d.xml";

        BackgroundWorker bw = new BackgroundWorker();
        BackgroundWorker bw1 = new BackgroundWorker();
        BackgroundWorker bw2 = new BackgroundWorker();
        BackgroundWorker bw3 = new BackgroundWorker();

        // 5 SEPT

        BackgroundWorker bw4 = new BackgroundWorker();
        BackgroundWorker bw5 = new BackgroundWorker();
        BackgroundWorker bw6 = new BackgroundWorker();

        // 28 - SEPT, On Request from LTC
        BackgroundWorker bw7 = new BackgroundWorker();

        //29 - AUG-2019, Sanitize XL - Post Upload
        BackgroundWorker bw8 = new BackgroundWorker();

        //10 - SEPT-2019, Sanitize XL - Post Clone
        BackgroundWorker bw9_PostClone = new BackgroundWorker();

        //18 - June-2020, Check Out Excel
        BackgroundWorker bw10_CheckOut = new BackgroundWorker();

        //18 - June-2020, Sanitize XL - Post Clone
        BackgroundWorker bw11_CheckIn = new BackgroundWorker();

        public Ribbon2d()
        {
            // Get a reference to the current assembly. This is where the ribbon XML is embedded.
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            // In this example, XML file must have a build action of "Embedded Resource".
            this.LoadXml(assembly, _embeddedResourceName);

            // TVS
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            //CTC
            bw1.DoWork += new DoWorkEventHandler(bw1_DoWork);
            bw1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw1_RunWorkerCompleted);

            //CTE
            bw2.DoWork += new DoWorkEventHandler(bw2_DoWork);
            bw2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw2_RunWorkerCompleted);

            //SyncTE
            bw3.DoWork += new DoWorkEventHandler(bw3_DoWork);
            bw3.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw3_RunWorkerCompleted);


            //REDESIGN TVS
            bw4.DoWork += new DoWorkEventHandler(bw4_DoWork);
            bw4.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw4_RunWorkerCompleted);


            //REDESIGN CTE
            bw5.DoWork += new DoWorkEventHandler(bw5_DoWork);
            bw5.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw5_RunWorkerCompleted);

            //REDESIGN CTC
            bw6.DoWork += new DoWorkEventHandler(bw6_DoWork);
            bw6.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw6_RunWorkerCompleted);

            //REDESIGN SyncTE - 28 SEPT
            bw7.DoWork += new DoWorkEventHandler(bw7_DoWork);
            bw7.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw7_RunWorkerCompleted);

            //Sanitize XL - Post Upload - 29 AUG
            bw8.DoWork += new DoWorkEventHandler(bw8_DoWork);
            bw8.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw8_RunWorkerCompleted);

            //Sanitize XL - Post Upload - 10-SEPT
            bw9_PostClone.DoWork += new DoWorkEventHandler(bw9_PostClone_DoWork);
            bw9_PostClone.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw9_PostClone_RunWorkerCompleted);

            //Check Out 18/06/2020
            bw10_CheckOut.DoWork += new DoWorkEventHandler(bw10_CheckOut_DoWork);
            bw10_CheckOut.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw10_CheckOut_RunWorkerCompleted);

            //Check In 18/06/2020
            bw11_CheckIn.DoWork += new DoWorkEventHandler(bw11_CheckIn_DoWork);
            bw11_CheckIn.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw11_CheckIn_RunWorkerCompleted);
        }

        private void bw11_CheckIn_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Check In Operation Completed. Please Refresh the Item Revision in Team-center and check if Excel Dataset is checked-In..");
            return;
        }

        private void bw11_CheckIn_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];
            try
            {
                CheckInOut.RunCheckInCheckOutMethod("CheckIn", application);
            }
            catch (Exception ex)
            {
                e.Result = "NOK";
                return;
            }
            e.Result = "OK";
        }

        private void bw10_CheckOut_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Check Out Operation Completed. Please Refresh the Item Revision in Team-center and check if Excel Dataset is checked-out..");
            return;
        }

        private void bw10_CheckOut_DoWork(object sender, DoWorkEventArgs e)
        {

            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];
            try
            {
                CheckInOut.RunCheckInCheckOutMethod("CheckOut", application);
            }
            catch (Exception ex)
            {
                e.Result = "NOK";
                return;
            }
            e.Result = "OK";

        }

        

        private void bw9_PostClone_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result == null || e.Result.ToString().Equals("NOK"))
            {
                MessageBox.Show("Sanitize XL - Post Clone Operation Failed. Please check the post clone log file for more details..");
                return;
            }

            MessageBox.Show("Sanitize XL - Post Clone Operation Completed. Please check the updated Excel dataset in Team-center..");
            return;
        }

        private void bw9_PostClone_DoWork(object sender, DoWorkEventArgs e)
        {

            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];
            bool parseFlag = false;
            try
            {
                parseFlag = RunSanitizeXL_PostClone_2(application);
            }
            catch (Exception ex)
            {
                e.Result = "NOK";
                return;
            }
            if (parseFlag == false)
            {
                e.Result = "NOK";
            }
            else
            {
                e.Result = "OK";
            }
        }

        private void bw8_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            return;
        }

        //Sanitize XL - 29 AUG
        private void bw8_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];

            using (var dialog = new SanitizeXL_PostTVS())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (application.ShowDialog(dialog) == DialogResult.OK)
                {
                }

            }
            e.Result = "OK";

        }

        private void bw7_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            return;
        }

        //Redesign SyncTE - 28 Sept
        private void bw7_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];

            using (var dialog = new SyncTEDialog())
            {
                if (application.ShowDialog(dialog) == DialogResult.OK)
                {
                }

            }
            e.Result = "OK";
        }

        private void bw6_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            return;
        }

        private void bw6_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];



            try
            {
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new System.Action(() =>
                {
                    CTCProgressDialog dialog = new CTCProgressDialog();
                    application.ShowDialog(dialog);
                }
                ));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



            e.Result = "OK";
        }

        private void bw5_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            return;
        }

        private void bw5_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];

            using (var dialog = new CTEProgressDialog())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (application.ShowDialog(dialog) == DialogResult.OK)
                {
                    using (var dialog1 = new MyCustomDialog4())
                    {

                        // The ShowDialog() extension method is exposed by:
                        // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                        if (application.ShowDialog(dialog1) == DialogResult.OK)
                        {
                        }
                    }
                }

            }
            e.Result = "OK";
        }

        private void bw4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            return;
        }

        private void bw4_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (SolidEdgeFramework.Application)genericlist[0];


            using (var dialog = new TVSProgressDialog())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (application.ShowDialog(dialog) == DialogResult.OK)
                {
                    //dialog.Close();
                    try
                    {
                        using (var dialog1 = new MyCustomDialog3())
                        {
                            //this.Close();
                            // The ShowDialog() extension method is exposed by:
                            // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                            if (application.ShowDialog(dialog1) == DialogResult.OK)
                            {
                                dialog1.Close();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("UI Initialization Issue: " + ex.Message);
                    }
                }
                e.Result = "OK";
            }
        }

        //SyncTE
        private void bw3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = e.Result as List<object>;
            if (genericlist == null || genericlist.Count == 0)
            {
                return;
            }
            String logFilePath = (String)genericlist[1];

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            
            //parse the log file and check for errors, use the start string to start parsing from
            String startString = "SE_SESSION - Initating Solid Edge Application";
            bool parseFlag = Utility.parseLog(logFilePath, startString);
            if (parseFlag == false)
            {
                MessageBox.Show("SyncTE Completed with Errors. To get more details open log file at "+ logFilePath);
                return;
            }
            MessageBox.Show("SyncTE Completed");
        }
        //SyncTE
        private void bw3_DoWork(object sender, DoWorkEventArgs e)
        {


            List<object> genericlist = e.Argument as List<object>;
            String xlFile = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            String assemblyFileName = (String)genericlist[2];
            if (System.IO.File.Exists(xlFile) == false)
            {
                Utlity.Log("File does not Exist: " + xlFile, logFilePath);
                e.Result = null;
                return;
            }

            if (System.IO.File.Exists(assemblyFileName) == false)
            {
                Utlity.Log("File does not Exist: " + assemblyFileName, logFilePath);
                e.Result = null;
                return;
            }

            if (System.IO.File.Exists(logFilePath) == false)
            {
                Utlity.Log("File does not Exist: " + logFilePath, logFilePath);
                e.Result = null;
                return;
            }

            try
            {
                ExcelData.readOccurenceVariablesFromTemplateExcel(xlFile, logFilePath);
                ExcelData.readOccurencePathFromTemplateExcel(xlFile, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
            }
            try
            {
                SolidEdgeInterface.SolidEdgeSync(assemblyFileName, logFilePath, "VALUE");
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
            }

            e.Result = genericlist;

        }

        //CTE
        private void bw2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Result as List<object>;
            String xlFile = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            application = (SolidEdgeFramework.Application)genericlist[2];

            List<Variable> AllVariablesList = ExcelData.getVariableDetails();

            Utlity.Log("variableList Size: " + AllVariablesList.Count, logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);

            using (var dialog = new MyCustomDialog4())
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }
        }
        //CTE
        private void bw2_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            String xlFile = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];

            application = (SolidEdgeFramework.Application)genericlist[2];
            try
            {
                Utlity.Log("readOccurenceVariablesFromTemplateExcel: ", logFilePath);
                ExcelData.readOccurenceVariablesFromTemplateExcel(xlFile, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("readOccurenceVariablesFromTemplateExcel: " + ex.Message, logFilePath);
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
                Utlity.Log("readOccurencePathFromTemplateExcel: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }

            try
            {
                Utlity.Log("readOccurencePathFromTemplateExcel: ", logFilePath);
                ExcelData.readOccurencePathFromTemplateExcel(xlFile, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("readOccurencePathFromTemplateExcel: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }

            try
            {
                if (SolidEdgeHighLighter.getOccurenceCount() == 0)
                {
                    Utlity.Log("--readOccurences-- ", logFilePath);
                    SolidEdgeHighLighter.readOccurences(logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("readOccurences: " + ex.Message, logFilePath);
                e.Result = null;
                return;

            }
            e.Result = genericlist;
        }

        //CTC
        private void bw1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = e.Result as List<object>;
            if (genericlist == null)
            {
                return;
            }
            if (genericlist.Count == 3)
            {
                String logFilePath = (String)genericlist[2];
                Utlity.Log("-----------------------------------------------------------------", logFilePath);
                Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
                MessageBox.Show("Custom Template Creation Completed");
            }
        }
        //CTC
        private void bw1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String folderToPublish = (String)genericlist[0];
            String assemblyFileName = (String)genericlist[1];
            String logFilePath = (String)genericlist[2];
            //SolidEdgeData.updateLinkedTemplate(assemblyFileName, logFilePath);
            SolidEdgeData1.copyLinkedDocumentsToPublishedFolder2(folderToPublish, assemblyFileName, logFilePath, true);
            //06SEPT - COMMENTED, AFTER CUSTOMER CLARIFICATION CAN UNCOMMENT
            //SolidEdgeOccurenceDelete.process(folderToPublish, assemblyFileName,logFilePath);

            // 01-OCT -- Search and Copy Drafts to the Custom Template Create Folder
            Utlity.Log("Copying Associated Drafts to Publish Folder: " + folderToPublish, logFilePath);
            String searchDrawingsFolder = System.IO.Path.GetDirectoryName(assemblyFileName);
            Utlity.Log("searchDrawingsFolder: " + searchDrawingsFolder, logFilePath);
            SolidEdgeData1.SearchAndcollectdrafts(assemblyFileName, folderToPublish, searchDrawingsFolder, logFilePath);
            //SolidEdgeRedefineLinks.ReplaceLinks(folderToPublish, logFilePath);

            try
            {
                // VariablePartsList - NULL, Suffix - ""
                ExcelComponentDeltaInterface.RenameComponentDetailsInMasterAssemblyTab(assemblyFileName, folderToPublish, null, "", logFilePath, "CTC");
            }
            catch (Exception ex)
            {
                Utlity.Log("RenameComponentDetailsInMasterAssemblyTab: " + ex.Message, logFilePath);
            }

            Utlity.Log("Custom Template Creation Completed " + System.DateTime.Now.ToString(), logFilePath);
            e.Result = genericlist;

        }

        //TVS
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Result as List<object>;
            String assemblyFileName = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            application = (SolidEdgeFramework.Application)genericlist[2];
            List<Variable> ALLVariablesList = SolidEdgeData1.getVariableDetails();

            Utlity.Log("variableList Size: " + ALLVariablesList.Count, logFilePath);
            Utlity.Log("featureList Size: " + SolidEdgeReadFeature.getFeatureLines().Count, logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);

            //application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }

            try
            {
                using (var dialog = new MyCustomDialog3())
                {
                    // The ShowDialog() extension method is exposed by:
                    // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                    if (application.ShowDialog(dialog) == DialogResult.OK)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("UI Initialization Issue: " + ex.Message, logFilePath);
            }
        }
        //TVS
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            String assemblyFileName = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            application = (SolidEdgeFramework.Application)genericlist[2];
            try
            {
                Utlity.Log("--readVariablesForEachOccurence-- ", logFilePath);
                SolidEdgeData1.readVariablesForEachOccurence(assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("readVariablesForEachOccurence: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }
            try
            {
                Utlity.Log("--traverseAssembly SolidEdgeData1-- ", logFilePath);
                SolidEdgeData1.traverseAssembly(assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SolidEdgeData1, traverseAssembly: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }
            try
            {
                Utlity.Log("--traverseAssembly SolidEgeData2-- ", logFilePath);
                SolidEgeData2.traverseAssembly(assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SolidEgeData2, traverseAssembly: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }


            try
            {
                //if (SolidEdgeHighLighter.getOccurenceCount() == 0)
                {
                    Utlity.Log("--readOccurences SolidEdgeHighLighter-- ", logFilePath);
                    SolidEdgeHighLighter.readOccurences(logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("readOccurences: " + ex.Message, logFilePath);
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
                Utlity.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }
            e.Result = genericlist;
        }

        public override void OnControlClick(RibbonControl control)
        {

            currentSESession = SolidEdgeCommunity.SolidEdgeUtils.Connect(false);

            currentSETCEObject = currentSESession.SolidEdgeTCE;
            if (currentSETCEObject != null)
            {
                currentSETCEObject.GetCurrentUserName(out userName);
                //MessageBox.Show("userName: " + userName);
            }
            int commandId = control.CommandId;
            //MessageBox.Show("commandId: " + commandId);
            //if (commandId == 11)
            //{
            //    UploadtoTCUsingSEEC(currentSETCEObject, currentSESession);
            //    DownloadXLfromTC_1(currentSETCEObject, currentSESession);
            //}

            //if (commandId != 3 & commandId != 8)
            //{
                //if (userName.Equals("") == false & loginFromSE.loggedInthroughUtility == true)
                //{
                    if (commandId == 0) // TVS
                    {
                        //MessageBox.Show("Inside TVS: " + commandId);
                        //if (loginFromSE.role.Equals("DBA", StringComparison.OrdinalIgnoreCase) == true)
                        ConnectToSolidEdge_v1();
                        //else
                           // MessageBox.Show("You have to be logged in as a DBA to use this function");

                    }
                    else if (commandId == 1) //CTE
                    {
                        //if (loginFromSE.role.Equals("DBA", StringComparison.OrdinalIgnoreCase) == true)
                            ConnectToTemplateExcel_v1();
                        //else
                            //MessageBox.Show("You have to be logged in as a DBA to use this function");
                    }
                    else if (commandId == 2) // Sanitize XL - post Add to TC
                    {
                        //if (loginFromSE.role.Equals("DBA", StringComparison.OrdinalIgnoreCase) == true)
                            RunSanitizeXL_Post_Add_To_TC_background();
                        //else
                            //MessageBox.Show("You have to be logged in as a DBA to use this function");
                    }
                    else if (commandId == 4) // Open
                    {
                        //if (loginFromSE.role.Equals("Designer", StringComparison.OrdinalIgnoreCase) == true)
                        OpenTheTemplateExcel();
                        //else
                        //    MessageBox.Show("You have to be logged in as a Designer to use this function");

                    }
                    else if (commandId == 5) // Sync TE
                    {
                        //if (loginFromSE.role.Equals("Designer", StringComparison.OrdinalIgnoreCase) == true)
                        SyncVariablesFromExcelToSolidEdge_v1();
                        //else
                        //    MessageBox.Show("You have to be logged in as a Designer to use this function");
                    }
                    else if (commandId == 6) // Sanitize XL - post clone
                    {
                        //if (loginFromSE.role.Equals("Designer", StringComparison.OrdinalIgnoreCase) == true)
                        RunSanitizeXL_PostClone_background();
                        //else
                        //    MessageBox.Show("You have to be logged in as a Designer to use this function");

                    }
                    else if (commandId == 7) //CTD
                    {
                        // if (loginFromSE.role.Equals("Designer", StringComparison.OrdinalIgnoreCase) == true)
                        CreateTemplateDerivative();
                        //else
                        //    MessageBox.Show("You have to be logged in as a Designer to use this function");
                    }
                    else if (commandId == 9)
                    {
                        if (loginFromSE.role.Equals("Designer", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            RunCheckOut_In_TC_background();
                        }
                        else
                            MessageBox.Show("You have to be logged in as a Designer to use this function");
                    }

                    else if (commandId == 10)
                    {
                        //   if (loginFromSE.role.Equals("Designer", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            RunCheckIn_In_TC_background();
                        }
                        //else
                        //    MessageBox.Show("You have to be logged in as a Designer to use this function");

                    }
                    

                    else
                    {
                        MessageBox.Show("Please login through the login utility and try again");
                    }
                //}
                //else
                //{

                //    currentSESession.DisplayAlerts = false;
                //    currentSETCEObject.GetCurrentUserName(out userName);

                //    if (userName.Equals(""))
                //    {
                //        if (commandId == 3) //Login as admin
                //        {
                //            loginFromSE.group = "dba";
                //            loginFromSE.role = "DBA";
                //            login();
                //        }
                //        else if (commandId == 8) //Login as Designer
                //        {
                //            loginFromSE.group = "Engineering";
                //            loginFromSE.role = "Designer";
                //            login();
                //        }
                //    }
                //    else
                //    {
                //        if (loginFromSE.loggedInthroughUtility == true)
                //        {

                //            login2nd();
                //            //MessageBox.Show("You are already logged in. Group or Role cannot be changed");
                //        }
                //        else
                //            MessageBox.Show("Login was not done through LogIn utility. Cannot continue");
                //    }
                //    currentSESession.DisplayAlerts = true;

                //}
            //}
        }

        private void DownloadXLfromTC(SolidEdgeTCE currentSETCEObject, SolidEdgeFramework.Application currentSESession)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)currentSESession.ActiveDocument;

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

            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "DownloadXLfromTC" + ".txt");
            SolidEdgeFramework.SolidEdgeTCE objSEEC = currentSESession.SolidEdgeTCE;

           

            SolidEdgeFramework.PropertySets propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
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
            System.Object vFileNames = null;
            int nFiles = 0;
            System.Object[] objArray = null;
            //Downloading all applicable drawings for current assembly
            foreach (KeyValuePair<string, string> pair in itemAndRevIds)
            {
                Utlity.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath);
                vFileNames = null;
                nFiles = 0;
                objSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                objArray = null;
                objArray = (System.Object[])vFileNames;
                foreach (System.Object o in objArray)
                {
                    string filename = (string)(o);
                    if (filename.Contains(".xlsx"))
                    {
                        System.Object[,] temp = new object[1, 1];
                        objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                            SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                            1, temp);
                        Utlity.Log("Downloaded XLSX of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                    }
                }
            }


        }


        private void DownloadXLfromTC_1(SolidEdgeTCE currentSETCEObject, SolidEdgeFramework.Application currentSESession)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)currentSESession.ActiveDocument;

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

            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "DownloadXLfromTC" + ".txt");
            SolidEdgeFramework.SolidEdgeTCE objSEEC = currentSESession.SolidEdgeTCE;

            SolidEdgeFramework.Application objApp = currentSESession;
            SolidEdgeFramework.SolidEdgeTCE ObjSEEC = currentSETCEObject;


            String cachePath = "";
            currentSETCEObject.GetPDMCachePath(out cachePath);

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
                
                dwExpandSelectionOptions =(int) SolidEdgeConstants.ExpandSelectionOptions.IncludeComponentsFromAssemblies;

                Utlity.Log("ExtractTranslatedFilesOfActiveDocument: " , logFilePath);
                ObjSEEC.ExtractTranslatedFilesOfActiveDocument(Location, (uint)dwExpandSelectionOptions, SEFiletypeFilters, RelationFilters, ReferanceFilters, ExportFileTypeFilter, out ListofFiles);
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
                ObjSEEC = null;
                objApp = null;
            }




        }

        public static SolidEdgeFramework.Application currentSESession = null;
        public static SolidEdgeFramework.SolidEdgeTCE currentSETCEObject = null;
        string userName = null;
        public void login()
        {
             
            loginFromSE loginObj = new loginFromSE();
            DialogResult dr = loginObj.ShowDialog();
            currentSETCEObject.SetTeamCenterMode(true);
            currentSETCEObject.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
            loginFromSE.loggedInthroughUtility = true;

        }

        public void login2nd()
        {
            // MessageBox.Show("call to login 2nd");

            string logfileDir = Utlity.CreateLogDirectory();
            String groupRoleLog = System.IO.Path.Combine(logfileDir, "LTC_FetchingGroupRoleFromTC.txt");
            Utility.Log("login2nd: ", groupRoleLog);

            loginFromSE loginObj = new loginFromSE();
            //for group rajesh code to fetch group from username
            loginFromSE.group = loginObj.GetGroupComboBoxText();
            loginFromSE.role = loginObj.GetRoleTextBoxText();
            DialogResult dr = loginObj.ShowDialog();
            currentSETCEObject.SetTeamCenterMode(true);
            Utility.Log("Group: "+ loginFromSE.group, groupRoleLog);
            Utility.Log("Role: " + loginFromSE.role, groupRoleLog);
            Utility.Log("userName: " + loginFromSE.userName, groupRoleLog);
            Utility.Log("password: " + loginFromSE.password, groupRoleLog);
            Utility.Log("URL: " + loginFromSE.URL, groupRoleLog);

            currentSETCEObject.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
            
            loginFromSE.loggedInthroughUtility = true;

        }

        public static void UploadtoTCUsingSEEC(SolidEdgeTCE currentSETCEObject, SolidEdgeFramework.Application currentSESession)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)currentSESession.ActiveDocument;

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
            //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "UploadtoTCUsingSEEC" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("UploadtoTCUsingSEEC Started @ " + System.DateTime.Now.ToString(), logFilePath);

            

            //String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            //String itemID = SEECAdaptor.getItemID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String fileName = System.IO.Path.GetFileName(assemblyFileName);
            Utlity.Log("UploadtoTCUsingSEEC: fileName: "+ fileName, logFilePath);
            String cachePath = "";
            currentSETCEObject.GetPDMCachePath(out cachePath);
            currentSETCEObject.SetTeamCenterMode(true);
            Utlity.Log("UploadtoTCUsingSEEC: cachePath: " + cachePath, logFilePath);
            SolidEdgeFramework.Application objApp = currentSESession;
            SolidEdgeFramework.SolidEdgeTCE ObjSEEC = currentSETCEObject;
            string strPDMCachePath = cachePath;
            string strInputFileName = fileName;
            string oldFileName = string.Empty;
            string strTranslatedFile = string.Empty;
            object[,] ListOfPropsForFileSaveAs = new object[6, 2];

            ListOfPropsForFileSaveAs[0, 0] = "Translation File Extension";
            ListOfPropsForFileSaveAs[0, 1] = "xlsx";
            ListOfPropsForFileSaveAs[1, 0] = "Dataset Type";
            ListOfPropsForFileSaveAs[1, 1] = "MSExcelX";
            ListOfPropsForFileSaveAs[2, 0] = "Named Reference Type";
            ListOfPropsForFileSaveAs[2, 1] = "excel";
            ListOfPropsForFileSaveAs[3, 0] = "Relation Name";
            ListOfPropsForFileSaveAs[3, 1] = "IMAN_specification";
            ListOfPropsForFileSaveAs[4, 0] = "Dataset Name";
            ListOfPropsForFileSaveAs[4, 1] = ""; //user can provide custom dataset name, default is input file dataset name
            ListOfPropsForFileSaveAs[5, 0] = "Dataset Description";
            ListOfPropsForFileSaveAs[5, 1] = "template xl"; //user can provide dataset description, default is empty


            oldFileName = Path.Combine(strPDMCachePath ,strInputFileName);
            strTranslatedFile = Path.Combine(strPDMCachePath, itemID + ".xlsx");
            Utlity.Log("UploadtoTCUsingSEEC strTranslatedFile.." + strTranslatedFile, logFilePath);
            if (File.Exists(strTranslatedFile) == true)
            {
                Utlity.Log("UploadtoTCUsingSEEC oldFileName.." + oldFileName, logFilePath);
                Utlity.Log("UploadtoTCUsingSEEC strTranslatedFile.." + strTranslatedFile, logFilePath);
                ObjSEEC.SetPDMPropsAndUploadTranslatedFile(oldFileName, ListOfPropsForFileSaveAs, strTranslatedFile);
            }else
            {
                Utlity.Log("UploadtoTCUsingSEEC strTranslatedFile does not Exist.." + strTranslatedFile, logFilePath);
            }

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("UploadtoTCUsingSEEC Ended @ " + System.DateTime.Now.ToString(), logFilePath);
        }

        private void  RunCheckOut_In_TC_background()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }
            SE_SESSION.setSolidEdgeSession(application);

            if (application.Documents.Count == 0)
            {
                MessageBox.Show("Open the Assembly that was Imported to TC..in Managed Mode");
                return;

            }

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw10_CheckOut.IsBusy != true)
            {
                bw10_CheckOut.RunWorkerAsync(arguments);
                if (bw10_CheckOut.IsBusy == true)
                    System.Windows.Forms.Application.DoEvents();
            }
            else
            {
                MessageBox.Show("RunCheckOut_In_TC_background is Already Running.Close the Old Process");
                return;
            }

        }

        private void RunCheckIn_In_TC_background()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }
            SE_SESSION.setSolidEdgeSession(application);

            if (application.Documents.Count == 0)
            {
                MessageBox.Show("Open the Assembly that was Imported to TC..in Managed Mode");
                return;

            }

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw11_CheckIn.IsBusy != true)
            {
                bw11_CheckIn.RunWorkerAsync(arguments);
                if (bw11_CheckIn.IsBusy == true)
                    System.Windows.Forms.Application.DoEvents();
            }
            else
            {
                MessageBox.Show("RunCheckIn_In_TC_background is Already Running.Close the Old Process");
                return;
            }

        }

        //private void RunSanitizeXL_PostClone_background()
        public void RunSanitizeXL_PostClone_background()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }
            SE_SESSION.setSolidEdgeSession(application);

            if (application.Documents.Count == 0)
            {
                MessageBox.Show("Open the Assembly that was Imported to TC..in Managed Mode");
                return;

            }

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw9_PostClone.IsBusy != true)
            {
                bw9_PostClone.RunWorkerAsync(arguments);
                if (bw9_PostClone.IsBusy == true)
                    System.Windows.Forms.Application.DoEvents();
            }
            else
            {
                MessageBox.Show("RunSanitizeXL_PostClone_background is Already Running.Close the Old Process");
                return;
            }
        }

        private void RunSanitizeXL_Post_Add_To_TC_background()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }
            SE_SESSION.setSolidEdgeSession(application);

            if (application.Documents.Count == 0)
            {
                MessageBox.Show("Open the Assembly that was Imported to TC..in Managed Mode");
                return;

            }

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw8.IsBusy != true)
            {
                bw8.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("RunSanitizeXL_Post_Add_To_TC is Already Running.Close the Old Process");
                return;
            }

        }



        //private void RunSanitizeXL_PostClone(SolidEdgeFramework.Application application)
        //{

        //    //SolidEdgeFramework.Application application = null;
        //    SolidEdgeDocument document = null;

        //    // Connect to running Solid Edge Instance
        //    //application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
        //    //if (application == null)
        //    //{
        //    //MessageBox.Show("Application is NULL");
        //    //return;
        //    //}

        //    SE_SESSION.setSolidEdgeSession(application);
        //    document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

        //    if (document == null)
        //    {
        //        MessageBox.Show("document is NULL");
        //        return;
        //    }
        //    String assemblyFileName = document.FullName;
        //    SolidEdgeData1.setAssemblyFileName(assemblyFileName);
        //    //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
        //    String stageDir = Utlity.CreateLogDirectory();
        //    String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "Sanitize_PostClone" + ".txt");

        //    Utlity.Log("-----------------------------------------------------------------", logFilePath, "CTD");
        //    Utlity.Log("Run Sanitize Excel Post Clone Utility Started @ " + System.DateTime.Now.ToString(), logFilePath, "CTD");

        //    Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath, "CTD");
        //    TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
        //    Utlity.Log("Initializing TC Services..", logFilePath, "CTD");
        //    TcAdaptor.TcAdaptor_Init(logFilePath);
        //    Utlity.Log("SEEC Login..", logFilePath, "CTD");
        //    SEECAdaptor.LoginToTeamcenter(logFilePath);

        //    String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
        //    String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
        //    String cachePath = SEECAdaptor.GetPDMCachePath();

        //    Utlity.Log("Download Excel file from Teamcenter..", logFilePath, "CTD");
        //    Utlity.Log("itemID.." + itemID, logFilePath, "CTD");
        //    Utlity.Log("RevID.." + RevID, logFilePath, "CTD");
        //    // Murali - 24-April 2020, clean up existing excel files in the cache... MUST NEEDED functionality before Sanitize XL is executed.

        //    String[] XlFiles = Directory.GetFiles(cachePath, "*", SearchOption.AllDirectories)
        //                                    .Select(path => Path.GetFullPath(path))
        //                                    .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
        //                                    .ToArray();
        //    foreach (String xLFile in XlFiles)
        //    {
        //        Utility.Log("Deleting xLFile: , " + xLFile, logFilePath);
        //        File.Delete(xLFile);
        //    }

        //    DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, cachePath, true, logFilePath, true);
        //    Utlity.Log("Sanitize XL..", logFilePath, "CTD");

        //    SanitizeXL_PostClone.read_all_items_in_cache(assemblyFileName, logFilePath);
        //    Utlity.Log("Clean up Dataset..", logFilePath, "CTD");
        //    TcAdaptor.PostCloneCleanUpExcelDataSet(itemID, RevID, logFilePath);

        //    Utlity.Log("Logout from TC..", logFilePath, "CTD");
        //    TcAdaptor.logout(logFilePath);

        //}


        private bool RunSanitizeXL_PostClone_2(SolidEdgeFramework.Application application)
        {

            //SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

      

            SE_SESSION.setSolidEdgeSession(application);
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

            if (document == null)
            {
                MessageBox.Show("document is NULL");
                return false;
            }
            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "Sanitize_PostClone" + ".txt");

            Utlity.Log("SEEC Login..", logFilePath);
            SEECAdaptor.LoginToTeamcenter(logFilePath);

            String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String cachePath = SEECAdaptor.GetPDMCachePath(logFilePath);
            if (cachePath == null || cachePath == "" || cachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                cachePath = Path.GetDirectoryName(assemblyFileName);
            }
            
            string bStrCurrentUser = null;
            SEECAdaptor.getSEECObject().GetCurrentUserName(out bStrCurrentUser);

            String password = bStrCurrentUser;

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Run Sanitize Excel Post Clone Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath);
            Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath);
            Utlity.Log("ID=" + bStrCurrentUser, logFilePath);
            Utlity.Log("Group=Engineering", logFilePath);
            Utlity.Log("Role=Designer", logFilePath);
            TcAdaptor.login(bStrCurrentUser, password, "Engineering", "Designer", logFilePath);
            Utlity.Log("Initializing TC Services..", logFilePath);
            TcAdaptor.get_Session_Log(logFilePath);
            TcAdaptor.TcAdaptor_Init(logFilePath);
            Utlity.Log("Download Excel file from Teamcenter..", logFilePath, "CTD");
            Utlity.Log("itemID.." + itemID, logFilePath);
            Utlity.Log("RevID.." + RevID, logFilePath);
            Utlity.Log("cachePath.." + cachePath, logFilePath);
            if (cachePath == null || cachePath == "" || cachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is Empty: So Sanitize is Aborted..", logFilePath);
                Utlity.Log("Logout from TC..", logFilePath, "CTD");
                TcAdaptor.logout(logFilePath);
            }
            // Murali - 24-April 2020, clean up existing excel files in the cache... MUST NEEDED functionality before Sanitize XL is executed.
            String[] XlFiles = Directory.GetFiles(cachePath, "*", SearchOption.AllDirectories)
                                            .Select(path => Path.GetFullPath(path))
                                            .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                                            .ToArray();

            Utlity.Log("XlFiles.." + XlFiles.Length, logFilePath);
            foreach (String xLFile in XlFiles)
            {
                Utility.Log("Deleting xLFile: , " + xLFile, logFilePath);
                File.Delete(xLFile);
            }
            Utlity.Log("RetrieveItemRevMOAndDownloadDatasetNR Start.." , logFilePath);
            DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, cachePath, true, logFilePath, true);
            Utlity.Log("Sanitize XL..", logFilePath, "CTD");

            SanitizeXL_PostClone.read_all_items_in_cache(assemblyFileName, logFilePath);
            Utlity.Log("Clean up Dataset..", logFilePath, "CTD");
            TcAdaptor.PostCloneCleanUpExcelDataSet(itemID, RevID, logFilePath);

            Utlity.Log("Logout from TC..", logFilePath, "CTD");
            TcAdaptor.logout(logFilePath);

            bool parseFlag = Utility.parseLog(logFilePath);
            return parseFlag;
        }

        // 15-10-2024 | Murali | Removed the SOA TC call, which will download the XL from TC.
        // 27-11-2024 | Murali | Not Able to Remove the SOA calls related to based_on property extraction.
        // 27-11-2024 | Murali | Not able to Remove the download template XL from TC to cache.
        //private void RunSanitizeXL_PostClone_1(SolidEdgeFramework.Application application)
        //{

        //    //SolidEdgeFramework.Application application = null;
        //    SolidEdgeDocument document = null;

        //    SE_SESSION.setSolidEdgeSession(application);
        //    document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

        //    if (document == null)
        //    {
        //        MessageBox.Show("document is NULL");
        //        return;
        //    }
        //    String assemblyFileName = document.FullName;
        //    SolidEdgeData1.setAssemblyFileName(assemblyFileName);
        //    //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
        //    String stageDir = Utlity.CreateLogDirectory();
        //    String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "Sanitize_PostClone" + ".txt");

        //    Utlity.Log("-----------------------------------------------------------------", logFilePath, "CTD");
        //    Utlity.Log("Run Sanitize Excel Post Clone Utility Started @ " + System.DateTime.Now.ToString(), logFilePath, "CTD");

        //    //Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath, "CTD");
           
        //    Utlity.Log("SEEC Login..", logFilePath, "CTD");
        //    SEECAdaptor.LoginToTeamcenter(logFilePath);

        //    String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
        //    String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
        //    String cachePath = SEECAdaptor.GetPDMCachePath();

        //    Utlity.Log("Download Excel file from Teamcenter..", logFilePath, "CTD");
        //    Utlity.Log("itemID.." + itemID, logFilePath, "CTD");
        //    Utlity.Log("RevID.." + RevID, logFilePath, "CTD");
        //    // Murali - 24-April 2020, clean up existing excel files in the cache... MUST NEEDED functionality before Sanitize XL is executed.

        //    String[] XlFiles = Directory.GetFiles(cachePath, "*", SearchOption.AllDirectories)
        //                                    .Select(path => Path.GetFullPath(path))
        //                                    .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
        //                                    .ToArray();
        //    foreach (String xLFile in XlFiles)
        //    {
        //        Utility.Log("Deleting xLFile: , " + xLFile, logFilePath);
        //        File.Delete(xLFile);
        //    }

        //    Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath, "CTD");
        //    Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath, "CTD");
        //    Utlity.Log("ID=dcproxy", logFilePath, "CTD");
        //    Utlity.Log("Group=Engineering", logFilePath, "CTD");
        //    Utlity.Log("Role=Designer", logFilePath, "CTD");
        //    TcAdaptor.login("dcproxy", "dcproxy", "Engineering", "Designer", logFilePath);
        //    Utlity.Log("Initializing TC Services..", logFilePath, "CTD");

        //    TcAdaptor.TcAdaptor_Init(logFilePath);

        //    //downloadExcelTemplateFromTeamcenter_1(application, document, logFilePath);
        //    DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, cachePath, true, logFilePath, true);
        //    Utlity.Log("Sanitize XL..", logFilePath, "CTD");

            
        //    SanitizeXL_PostClone.read_all_items_in_cache_SEEC(assemblyFileName, logFilePath);

        //    Utlity.Log("Logout TC SOA Services..", logFilePath, "CTD");
        //    TcAdaptor.logout(logFilePath);

        //    Utlity.Log("Upload To TC Using SEEC", logFilePath, "CTD");
        //    UploadtoTCUsingSEEC(currentSETCEObject, currentSESession);

        //}

        private void SyncVariablesFromExcelToSolidEdge_v1()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }

            SE_SESSION.setSolidEdgeSession(application);
            SolidEdgeDocument document = null;
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

            if (document == null)
            {
                MessageBox.Show("document is NULL");
                return;
            }

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw7.IsBusy != true)
            {
                bw7.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("SyncTE is Already Running.Close the Old Process");
                return;
            }
        }

        private void CustomTemplateCreate_v1()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }

            SE_SESSION.setSolidEdgeSession(application);


            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw6.IsBusy != true)
            {
                bw6.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("CTC is Already Running.Close the Old Process");
                return;
            }

        }

        private void ConnectToSolidEdge_v1()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }
            SE_SESSION.setSolidEdgeSession(application);

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw4.IsBusy != true)
            {
                bw4.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("TVS is Already Running.Close the Old Process");
                return;
            }

        }

        public static TemplateCreationWizard tcw = null;
        private void CreateTemplateDerivative()
        {

            tcw = new TemplateCreationWizard();
            bool tc_mode = false;
            SolidEdgeFramework.Application application = SolidEdgeCommunity.SolidEdgeUtils.Connect(false);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, "CreateTemplateDerivative.txt");



            try
            {
                SolidEdgeFramework.Application currentSESession = SolidEdgeCommunity.SolidEdgeUtils.Connect(false);
                SolidEdgeFramework.SolidEdgeTCE currentSETCEObject = currentSESession.SolidEdgeTCE;
                string userName = null;
                currentSETCEObject.GetCurrentUserName(out userName);
                if (userName.Equals(""))
                    tc_mode = false;
                else
                    tc_mode = true;
                if (currentSESession.Documents.Count != 0)
                {
                    MessageBox.Show("Close all open documents and try again");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }


            if (tc_mode == true)
            {
                if (tcw.ShowDialog() == DialogResult.OK)
                {

                }
            }
            else
                MessageBox.Show("Please log-in in to Teamcenter and try again");

        }



        private void Duplicate()
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeFramework.SolidEdgeDocument document = null;

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


            using (var dialog = new MyCustomDialog2(assemblyFileName))
            {
                // The ShowDialog() extension method is exposed by:
                // using SolidEdgeFramework.Extensions (SolidEdge.Community.dll)
                if (application.ShowDialog(dialog) == DialogResult.OK)
                {
                }
            }


        }



        private void CustomTemplateCreate()
        {
            String folderToPublish = OpenTemplatePublishDialog();
            if (folderToPublish == null || folderToPublish.Equals("") == true)
            {
                return;
            }

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
            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "CTC" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            if (Utlity.checkifFilesAlreadyInFolderToPublish(folderToPublish, logFilePath) == true)
            {
                MessageBox.Show("Delete All Files in " + folderToPublish + " To Proceed Further");
                return;
            }

            List<object> arguments = new List<object>();
            arguments.Add(folderToPublish);
            arguments.Add(assemblyFileName);
            arguments.Add(logFilePath);

            if (bw1.IsBusy != true)
            {
                bw1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker Thread is Busy, CTC", logFilePath);
            }

        }

        // Sync the Variables From Excel To SolidEdge
        /*private void SyncVariablesFromExcelToSolidEdge()
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
            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            String AssemstageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "SyncTE" + ".txt");

            String xlFile = System.IO.Path.Combine(AssemstageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            List<object> arguments = new List<object>();
            arguments.Add(xlFile);
            arguments.Add(logFilePath);
            arguments.Add(assemblyFileName);

            if (bw3.IsBusy != true)
            {
                bw3.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker is Busy, SyncTE", logFilePath);
            }

        }*/

        private void OpenTheTemplateExcel()
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
            //SolidEdgeData.setAssemblyFileName(assemblyFileName);
            String assemDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String xlFile = System.IO.Path.Combine(assemDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "OpenTemplateXL" + ".txt");

            //String xlFile = System.IO.Path.Combine(assemDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            if (System.IO.File.Exists(xlFile) == true)
            {
                Utlity.Log("xlFile will be deleted and redownloaded..", logFilePath);
                //OpenTemplateExcelOptionDialog(assemDir);
                File.Delete(xlFile);
            }
            //else
            //{

            //MessageBox.Show("Could Not Find the Template Excel, Downloading from Team-center...");
             downloadExcelTemplateFromTeamcenter(assemblyFileName, logFilePath);
            //downloadExcelTemplateFromTeamcenter_1(application, document, logFilePath);
            if (System.IO.File.Exists(xlFile) == true)
            {
                OpenTemplateExcelOptionDialog(assemDir);
            }
            else
            {
                MessageBox.Show("Template Excel for Assembly: " + assemblyFileName + ": does not Exist");
            }
            //}

        }

        // 14-10-2024 - MURALI - Removed SOA call.
        public static void downloadExcelTemplateFromTeamcenter_1(SolidEdgeFramework.Application application, SolidEdgeDocument document, string logFilePath)
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
            SolidEdgeFramework.SolidEdgeTCE objSEEC = currentSESession.SolidEdgeTCE;

            if (objSEEC == null)
            {
                MessageBox.Show("objSEEC is NULL");
                return;
            }

            SolidEdgeFramework.Application objApp = currentSESession;
            SolidEdgeFramework.SolidEdgeTCE ObjSEEC = currentSETCEObject;

            if (objSEEC == null)
            {
                MessageBox.Show("objSEEC is NULL");
                return;
            }

            String cachePath = "";
            currentSETCEObject.GetPDMCachePath(out cachePath);

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
                
                ObjSEEC.ExtractTranslatedFilesOfActiveDocument(Location, (uint)dwExpandSelectionOptions, SEFiletypeFilters, RelationFilters, ReferanceFilters, ExportFileTypeFilter, out ListofFiles);

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

        // Designer Functionality Only
        //public static void downloadExcelTemplateFromTeamcenter_SOA_working(string assemblyFileName, String logFilePath)
        //{
        //    //TcAdaptor Tc = new TcAdaptor();
        //    bool logIn_Success = TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
        //    //TcAdaptor.TcAdaptor_Init(logFilePath);
        //    SEECAdaptor.LoginToTeamcenter(logFilePath);

        //    String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
        //    String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
        //    String stageDir = SEECAdaptor.GetPDMCachePath();

        //    //checkout

        //    SolidEdgeFramework.Application application = SEECAdaptor.GetApplication;
        //    if (application == null)
        //    {
        //        MessageBox.Show("Solid Edge Application is NULL");
        //        return;
        //    }
        //    try
        //    {
        //        Utility.Log("RunCheckInCheckOutMethod", logFilePath);
        //       // CheckInOut.RunCheckInCheckOutMethod("CheckOut", application);
        //       // {
        //            SolidEdgeData1.setAssemblyFileName(assemblyFileName);

        //            Utlity.Log("-----------------------------------------------------------------", logFilePath);
        //            Utlity.Log("Run " + "Checkout" + "  Started @ " + System.DateTime.Now.ToString(), logFilePath);

        //            Utility.Log(" Excel: getItemRevisionQuery..", logFilePath);
        //            ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID, logFilePath);

        //            if (itemRevMO == null)
        //            {
        //                Utility.Log( " Excel: item REV model Object is NULL/Empty", logFilePath);
        //                return;

        //            }

        //            Utility.Log(" Excel: isDataSetAvailable..", logFilePath);
        //            ModelObject dataSetMo = TcAdaptor.isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");

        //            if (dataSetMo == null)
        //            {
        //                Utility.Log(" Excel:..Excel Dataset is not available under the Item Revision..", logFilePath);
        //                return;
        //            }
             
        //            if (CheckInOut.isExcelDSCheckedOut(dataSetMo, logFilePath))
        //            {
        //                Utility.Log("INFO : Excel Dataset is already Checked Out: ", logFilePath);
        //                //return;
        //            }

        //            else
        //            {
        //                ModelObject dsMo = TcAdaptor.checkOutModelObject(dataSetMo, logFilePath);
        //                if (dsMo != null)
        //                {
        //                    Utility.Log("Excel Dataset is Checked Out...", logFilePath);
        //                    //return;
        //                }
        //            }                   
        //       // }
        //    }

        //    catch (Exception ex)
        //    {
        //        Utility.Log("Exception in RunCheckInCheckOutMethod :" + ex.ToString(), logFilePath);
        //        //e.Result = "NOK";
        //        //return;
        //    }

        //    DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, stageDir, true, logFilePath, true);
        //    TcAdaptor.logout(logFilePath);
        //}


        public static void downloadExcelTemplateFromTeamcenter(string assemblyFileName, String logFilePath)
        {
            SEECAdaptor.LoginToTeamcenter(logFilePath);
            string bStrCurrentUser = null;
            SEECAdaptor.getSEECObject().GetCurrentUserName(out bStrCurrentUser);

            String password = bStrCurrentUser;

            Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath, "CTD");
            Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath, "CTD");
            Utlity.Log("ID=" + bStrCurrentUser, logFilePath, "CTD");
            Utlity.Log("Group=Engineering", logFilePath, "CTD");
            Utlity.Log("Role=Designer", logFilePath, "CTD");
            TcAdaptor.login(bStrCurrentUser, password, "Engineering", "Designer", logFilePath);
            Utlity.Log("Initializing TC Services..", logFilePath, "CTD");

            TcAdaptor.TcAdaptor_Init(logFilePath);
           

            String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String stageDir = SEECAdaptor.GetPDMCachePath();

            if (stageDir == null || stageDir == "" || stageDir.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                stageDir = Path.GetDirectoryName(assemblyFileName);
            }

            //checkout

            SolidEdgeFramework.Application application = SEECAdaptor.GetApplication;
            if (application == null)
            {
                MessageBox.Show("Solid Edge Application is NULL");
                return;
            }
            try
            {
                Utility.Log("RunCheckInCheckOutMethod", logFilePath);
                // CheckInOut.RunCheckInCheckOutMethod("CheckOut", application);
                // {
                SolidEdgeData1.setAssemblyFileName(assemblyFileName);

                Utlity.Log("-----------------------------------------------------------------", logFilePath);
                Utlity.Log("Run " + "Checkout" + "  Started @ " + System.DateTime.Now.ToString(), logFilePath);

                Utility.Log(" Excel: getItemRevisionQuery..", logFilePath);
                ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID, logFilePath);

                if (itemRevMO == null)
                {
                    Utility.Log(" Excel: item REV model Object is NULL/Empty", logFilePath);
                    return;

                }

                Utility.Log(" Excel: isDataSetAvailable..", logFilePath);
                ModelObject dataSetMo = TcAdaptor.isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");

                if (dataSetMo == null)
                {
                    Utility.Log(" Excel:..Excel Dataset is not available under the Item Revision..", logFilePath);
                    return;
                }

                if (CheckInOut.isExcelDSCheckedOut(dataSetMo, logFilePath))
                {
                    Utility.Log("INFO : Excel Dataset is already Checked Out: ", logFilePath);
                    //return;
                }

                else
                {
                    // 29-11-2024 | Murali | CHECK OUT IS NOT NEEDED
                    // 05-12-2024 | Murali | CHECK OUT NEEDED
                    ModelObject dsMo = TcAdaptor.checkOutModelObject(dataSetMo, logFilePath);
                    if (dsMo != null)
                    {
                        Utility.Log("Excel Dataset is Checked Out...", logFilePath);
                        //return;
                    }
                    // 29-11-2024 | Murali | CHECK OUT IS NOT NEEDED
                }
                // }
            }

            catch (Exception ex)
            {
                Utility.Log("Exception in RunCheckInCheckOutMethod :" + ex.ToString(), logFilePath);
                //e.Result = "NOK";
                //return;
            }

            DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, stageDir, true, logFilePath, true);
            TcAdaptor.logout(logFilePath);
        }

        //20-01-20 Methun
        public static bool checkForExcelTemplateInTeamcenter(string assemblyFileName, String logFilePath)
        {
            //TcAdaptor Tc = new TcAdaptor();
            bool logIn_Success = TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
            //TcAdaptor.TcAdaptor_Init(logFilePath);
            SEECAdaptor.LoginToTeamcenter(logFilePath);

            String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String stageDir = SEECAdaptor.GetPDMCachePath();

            bool excelAvlbl = DownloadDatasetNamedReference.checkForExcelDataset(itemID, RevID, stageDir, true, logFilePath, true);
            TcAdaptor.logout(logFilePath);
            return excelAvlbl;
        }

        private String OpenTemplatePublishDialog()
        {
            String folderToPublish = "";
            System.Windows.Forms.FolderBrowserDialog FD = new System.Windows.Forms.FolderBrowserDialog();
            //FD.Filter = "Excel Files(*.xlsx)|*.xlsx";
            FD.Description = "Select Folder to Publish Template";
            if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderToPublish = FD.SelectedPath;
            }
            return folderToPublish;

        }


        private void OpenTemplateExcelOptionDialog(String stageDir)
        {
            System.Windows.Forms.OpenFileDialog FD = new System.Windows.Forms.OpenFileDialog();
            FD.Filter = "Excel Files(*.xlsx)|*.xlsx";
            FD.InitialDirectory = stageDir;
            if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileToOpen = FD.FileName;

                System.IO.FileInfo File = new System.IO.FileInfo(FD.FileName);
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                process.StartInfo = startInfo;
                startInfo.FileName = File.FullName;
                process.Start();
            }

        }

        private void ConnectToTemplateExcel_v1()
        {
            SolidEdgeFramework.Application application = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Application is NULL");
                return;
            }
            SE_SESSION.setSolidEdgeSession(application);

            List<object> arguments = new List<object>();
            arguments.Add(application);

            if (bw5.IsBusy != true)
            {
                bw5.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("CTE is Already Running.Close the Old Process");
                return;
            }



        }

        /*private void ConnectToTemplateExcel()
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
            String XLStageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "CTE" + ".txt");

            String xlFile = System.IO.Path.Combine(XLStageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            if (System.IO.File.Exists(xlFile) == false)
            {
                MessageBox.Show(xlFile + " is Missing.");
                Utlity.Log(xlFile + " is Missing.", logFilePath);
                return;
            }

            List<object> arguments = new List<object>();
            arguments.Add(xlFile);
            arguments.Add(logFilePath);
            arguments.Add(application);

            if (bw2.IsBusy != true)
            {
                bw2.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker Thread is Busy, CTE", logFilePath);
            }

        }*/

        [STAThread]
        private void ConnectToSolidEdge()
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
            if (System.IO.File.Exists(xlFileName) == true)
            {
                MessageBox.Show("Template is Already Published. Delete it To Use TVS");
                return;
            }

            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "TVS" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            List<object> arguments = new List<object>();
            arguments.Add(assemblyFileName);
            arguments.Add(logFilePath);
            arguments.Add(application);

            if (bw.IsBusy != true)
            {
                bw.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker Thread is Busy, TVS", logFilePath);
            }
        }
    }
}
