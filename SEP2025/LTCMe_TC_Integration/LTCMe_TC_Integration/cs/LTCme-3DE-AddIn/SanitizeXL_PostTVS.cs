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
using Creo_TC_Live_Integration.TcDataManagement;
using DemoAddInTC.controller;
using DemoAddInTC.se;
using DemoAddInTC.utils;
using SolidEdgeFramework;

namespace DemoAddInTC
{
    public partial class SanitizeXL_PostTVS : Form
    {
        public SanitizeXL_PostTVS()
        {
            
            InitializeComponent();
            this.progressBar1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String exceltoUpload = tmpPublishFolder.Text;

            if (File.Exists(exceltoUpload) == false)
            {               
                MessageBox.Show("Excel file does not exist");
                return;
            }

            List<object> arguments = new List<object>();
            arguments.Add(exceltoUpload);

            if (backgroundWorker1.IsBusy != true)
            {
                this.progressBar1.Visible = true;
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("SanitizeXL_PostUpload_To_TC is Already Running.Close the Old Process");
                return;
            } 

            
        }

        private void RunSanitizeXL_PostUpload_To_TC(String excelToUpload)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

            // Connect to running Solid Edge Instance
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                return;
            }

            SE_SESSION.setSolidEdgeSession(application);
            if (application.Documents.Count != 0)
            {
                document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            }

            if (document == null)
            {                
                return;
            }
            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "Sanitize_PostTVS" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            //TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
            //TcAdaptor.TcAdaptor_Init(logFilePath);
            SEECAdaptor.LoginToTeamcenter(logFilePath);

            String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String cachePath = SEECAdaptor.GetPDMCachePath();

            if (cachePath == null || cachePath == "" || cachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                cachePath = Path.GetDirectoryName(assemblyFileName);
            }

            String cacheExcel = Path.Combine(cachePath,Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            Utlity.Log("Copying XL to cache.." + cachePath, logFilePath);
            // copy the XL to cache path, rename the XL and upload to TC.
            if (File.Exists(cacheExcel) == false)
            {
                try
                {
                    File.Copy(excelToUpload, cacheExcel);
                }
                catch (Exception ex)
                {
                    Utlity.Log("Copying XL to cache.. Failed" + ex.Message, logFilePath);
                    return;
                }
            }
            else
            {
                try
                {
                    File.Delete(cacheExcel);
                }
                catch (Exception ex)
                {
                    Utlity.Log("Delete XL In cache.. Failed" + ex.Message, logFilePath);
                    return;
                }
                try
                {
                    File.Copy(excelToUpload, cacheExcel);
                }
                catch (Exception ex)
                {
                    Utlity.Log("Copying XL to cache.. Failed" + ex.Message, logFilePath);
                    return;
                }

            }

            //Utlity.Log("Uploading excel to TC.." + cacheExcel, logFilePath);
            //TcAdaptor.uploadExcelToTC(cacheExcel, logFilePath);
            //Utlity.Log("Upload completed to Teamcenter..", logFilePath);
            Utlity.Log("Sanitize the XL ..", logFilePath);
            Utility.Log("connecting.....",logFilePath);
            //DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemID, RevID, cachePath, true, logFilePath, true);

           // TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
            SEECAdaptor.LoginToTeamcenter(logFilePath);
            string bstruser = null;
            Utility.Log("Connecting to SE....", logFilePath);
            SEECAdaptor.getSEECObject().GetCurrentUserName(out bstruser);
            Utility.Log("Getting username...", logFilePath);
            string password = bstruser;
            try
            {
                TcAdaptor.login(bstruser, password, "DBA", "dba", logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("exception while log in:" + ex.Message, logFilePath);
                Utility.Log("exception" + ex.StackTrace, logFilePath);
            }
            Utility.Log("Connecting to Teamcenter with user: " + bstruser, logFilePath);
            Utility.Log("passwrod:" + password, logFilePath);
            TcAdaptor.TcAdaptor_Init(logFilePath);
            
            SanitizeXL_PostUpload_Logic.read_all_items_in_cache(assemblyFileName, logFilePath);
            //Utlity.Log("Sanitize the XL .. completed..", logFilePath);
            Utlity.Log("Uploading excel to TC.." + cacheExcel, logFilePath);
            //changing again
            TcAdaptor.uploadExcelToTC(bstruser, password, "", "", cacheExcel, logFilePath);
            TcAdaptor.logout(logFilePath);

        }

        private void OpenTemplateExcelOptionDialog()
        {
            System.Windows.Forms.OpenFileDialog FD = new System.Windows.Forms.OpenFileDialog();
            FD.Filter = "Excel Files(*.xlsx)|*.xlsx";
           
            if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string folderToPublish = FD.FileName;

                if (folderToPublish == null || folderToPublish.Equals("") == true)
                {
                    return;
                }
                tmpPublishFolder.Text = folderToPublish;
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            String folderToPublish = "";
            Thread worker = new Thread(() =>
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                //dialog.ShowHelp = true;
                DialogResult result = dialog.ShowDialog();
                folderToPublish = dialog.FileName;
                //folderToPublish = this.tmpPublishFolder.Text;
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
            worker.Join();

            if (folderToPublish == null || folderToPublish.Equals("") == true)
            {
                return;
            }
            tmpPublishFolder.Text = folderToPublish;
            //OpenTemplateExcelOptionDialog();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
            SolidEdgeFramework.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            String excelFiletoUpload = (String)genericlist[0];
            RunSanitizeXL_PostUpload_To_TC(excelFiletoUpload);
            
            e.Result = "OK";

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.progressBar1.Visible = false;
            MessageBox.Show("Sanitized Excel Template Uploaded to Teamcenter..");
            this.Close();
           
            return;
        }
    }
}
