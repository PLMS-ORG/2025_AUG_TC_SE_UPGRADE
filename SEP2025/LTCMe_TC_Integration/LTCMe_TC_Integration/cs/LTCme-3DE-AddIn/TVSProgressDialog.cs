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

namespace DemoAddInTC
{
    public partial class TVSProgressDialog : Form
    {
        public TVSProgressDialog()
        {
            InitializeComponent();
            String fileName = getSolidEdgeAssemblyFileName();

            if (fileName == null || fileName.Equals("") == true)
            {
                this.label2.Text = "No Active Assembly File Available in Solid Edge";
            }
            else
            {
                bool eligible = checkFileEligibility(fileName);
                String AsmFileName = System.IO.Path.GetFileName(fileName);
                this.label2.Text = AsmFileName;

                if (eligible == false)
                {
                    MessageBox.Show("TVS can be Invoked Only on ASM files");
                    this.button1.Enabled = false;
                }
            }
        }

        // Read Assembly
        private void button1_Click_1(object sender, EventArgs e)
        {
            this.progressBar1.Visible = true;
            this.button1.Enabled = false;
            ConnectToSolidEdge();
        }

        private String getSolidEdgeAssemblyFileName()
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
            if (application != null && application.ActiveDocument != null)
            {
                document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            }

            if (document == null)
            {
                this.progressBar1.Visible = false;
                MessageBox.Show("document is NULL");
                return "";
            }
            //MessageBox.Show(document.FullName);

            fileName = document.FullName;
            return fileName;

        }

        private bool checkFileEligibility(string FullName)
        {
            String fileName = System.IO.Path.GetFileName(FullName);
            if (fileName.EndsWith(".asm") == true)
            {
                return true;
            }
            else
            {
                return false;
            }

        }


        [STAThread]
        private void ConnectToSolidEdge()
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

            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "TVS" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);


            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            String AssemblyDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String xlFileName = System.IO.Path.ChangeExtension(assemblyFileName, "xlsx");

            if (System.IO.File.Exists(xlFileName) == true)
            {
                Utlity.Log("Deleting the cache file", logFilePath);
                File.Delete(xlFileName);
            }
            // 26-11-2024 | Murali | check excel in TC is done through SEEC and not through SOA API
            //20-01-20 Methun
            //bool excelAvlblInTeamcenter = Ribbon2d.checkForExcelTemplateInTeamcenter(assemblyFileName, logFilePath);
            Ribbon2d.downloadExcelTemplateFromTeamcenter_1(application, document, logFilePath);

            bool excelAvlblInTeamcenter = false;
            if (System.IO.File.Exists(xlFileName) == true)
            {
                excelAvlblInTeamcenter = true;
            }
            
            if (excelAvlblInTeamcenter == true)
            {
                this.progressBar1.Visible = false;
                Utlity.Log("Excel is already available in Teamcenter.", logFilePath);
                MessageBox.Show("Template is already available in Teamcenter. Delete it To Use TVS");
                return;
            }
            //20-01-20 Methun
            // 26-11-2024 | Murali | check excel in TC is done through SEEC and not through SOA API

            List<object> arguments = new List<object>();
            arguments.Add(assemblyFileName);
            arguments.Add(logFilePath);
            arguments.Add(application);

            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker Thread is Busy, TVS", logFilePath);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
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
                this.DialogResult = DialogResult.Cancel;
                // this.Dispose();
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
                this.DialogResult = DialogResult.Cancel;
                //this.Dispose();
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
                this.DialogResult = DialogResult.Cancel;
                //this.Dispose();
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
                this.DialogResult = DialogResult.Cancel;
                //this.Dispose();
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
                this.DialogResult = DialogResult.Cancel;

                //this.Dispose();
                Utlity.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }
            e.Result = genericlist;

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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

            this.DialogResult = DialogResult.OK;

            //application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            //if (application == null)
            //{
            //    MessageBox.Show("Application is NULL");
            //    return;
            //}
            //application.Activate();            

        }

        private void TVSProgressDialog_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            About a = new About();
            a.ShowDialog();
        }


    }
}
