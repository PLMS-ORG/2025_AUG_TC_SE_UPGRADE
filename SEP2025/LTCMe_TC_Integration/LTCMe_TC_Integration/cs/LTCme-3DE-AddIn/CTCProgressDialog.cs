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

namespace DemoAddInTC
{
    public partial class CTCProgressDialog : Form
    {
       
        public CTCProgressDialog()
        {
            InitializeComponent();

        }  

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String folderToPublish = (String)genericlist[0];
            String assemblyFileName = (String)genericlist[1];
            String logFilePath = (String)genericlist[2];
            //SolidEdgeData.updateLinkedTemplate(assemblyFileName, logFilePath);
            SolidEdgeData1.copyLinkedDocumentsToPublishedFolder2(folderToPublish, assemblyFileName, logFilePath, true);

            SolidEdgeOccurenceDelete.process(folderToPublish, assemblyFileName, logFilePath);
            Utlity.Log("Custom Template Creation Completed " + System.DateTime.Now.ToString(), logFilePath);
            e.Result = genericlist;
        }
        

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = e.Result as List<object>;
            if (genericlist == null)
            {
                return;
            }
            if (genericlist.Count == 3)
            {
                this.DialogResult = DialogResult.OK;
                String logFilePath = (String)genericlist[2];
                Utlity.Log("-----------------------------------------------------------------", logFilePath);
                Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
                MessageBox.Show("Custom Template Creation Completed");
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }

           

        }

        private void CTCProgressDialog_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            About a = new About();
            a.ShowDialog();
        }

        
        private void Browse_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {                
                this.textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void Create_Click(object sender, EventArgs e)
        {
            CustomTemplateCreate();
        }


        private void CustomTemplateCreate()
        {
            String folderToPublish = this.textBox1.Text;
            if (folderToPublish == null || folderToPublish.Equals("") == true)
            {
                MessageBox.Show("Select the Folder to Publish");
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

            if (backgroundWorker1.IsBusy != true)
            {
                this.progressBar1.Visible = true;
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker Thread is Busy, CTC", logFilePath);
            }

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

       
    }
}
