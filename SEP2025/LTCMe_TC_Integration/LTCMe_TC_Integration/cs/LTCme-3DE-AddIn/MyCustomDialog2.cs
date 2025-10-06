using DemoAddInTC.controller;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC
{
    public partial class MyCustomDialog2 : Form
    {
        String m_assemblyFileName = "";
       

        public MyCustomDialog2(String assemblyFileName)
        {
            InitializeComponent();
            m_assemblyFileName = assemblyFileName;

            

        }

        // OK
        private void button1_Click(object sender, EventArgs e)
        {
            bool linkedDrawingsOption = true;
            bool CopyMasterExcelOption = true;

            if (radioButton1.Checked == true)
            {
                linkedDrawingsOption = true;
            }
            else if (radioButton2.Checked == true)
            {
                linkedDrawingsOption = false;
            }

            if (radioButton3.Checked == true)
            {
                CopyMasterExcelOption = true;
            }
            else if (radioButton4.Checked == true)
            {
                CopyMasterExcelOption = false;
            }
            String DuplicateFolderPath = this.textBox1.Text;
            String searchDrawingsFolder = this.textBox2.Text;
            if (DuplicateFolderPath == null || DuplicateFolderPath.Equals("") == true)
            {
                MessageBox.Show("Select the Directory to Duplicate", "Duplicate Folder");
                return;
            }
            string[] filePaths = Directory.GetFiles(DuplicateFolderPath, "*.*",
                                         SearchOption.AllDirectories);

            if (filePaths != null && filePaths.Length > 0)
            {
                MessageBox.Show("Clear the Directory " + DuplicateFolderPath + " And Proceed To Duplicate");
                return;
            }

            if (linkedDrawingsOption == true)
            {
                if (searchDrawingsFolder == null || searchDrawingsFolder.Equals("") == true)
                {
                    MessageBox.Show("Search Drawings Folder is Empty", "Search Drawings Folder");
                    return;
                }
            }

            //this.Dispose();
            SolidEdgeData1.setAssemblyFileName(m_assemblyFileName);
            // 15-OCT change As per LTC Request,
                //String xlFilePath = System.IO.Path.ChangeExtension(m_assemblyFileName, ".xlsx");
                //if (xlFilePath == null || xlFilePath.Equals("") == true)
                //{
                //    MessageBox.Show("Template is Missing in the Assembly");
                //    return;
                //}
                //if (System.IO.File.Exists(xlFilePath) == false)
                //{
                //    MessageBox.Show("Template is Missing in the Assembly");
                //    return;
                //}
            

            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(m_assemblyFileName) + "_" + "Duplicate" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            List<object> arguments = new List<object>();
            arguments.Add(DuplicateFolderPath);
            arguments.Add(m_assemblyFileName);
            arguments.Add(logFilePath);
            arguments.Add(CopyMasterExcelOption);
            arguments.Add(linkedDrawingsOption);
            arguments.Add(searchDrawingsFolder);

            if (backgroundWorker1.IsBusy != true)
            {
                this.progressBar1.Visible = true;
                backgroundWorker1.RunWorkerAsync(arguments);
            }

            

        }

        //Browse - Duplicate Folder
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                //
                // The user selected a folder and pressed the OK button.                
                //                
                //MessageBox.Show("Selected Directory to Duplicate To: " + folderBrowserDialog1.SelectedPath, "Duplicate Folder To");
                this.textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        

        //Browse - Search Drawing Folder
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                //
                // The user selected a folder and pressed the OK button.                
                //                
                //MessageBox.Show("Selected Directory to Search Drawings: " + folderBrowserDialog1.SelectedPath, "Search Drawings @");
                this.textBox2.Text = folderBrowserDialog1.SelectedPath;
            }


        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String DuplicateFolderPath = (String)genericlist[0];
            String m_assemblyFileName = (String)genericlist[1];
            String logFilePath = (String)genericlist[2];
            bool CopyMasterExcelOption = (bool)genericlist[3];
            bool linkedDrawingsOption = (bool)genericlist[4];
            String searchDrawingsFolder = (String)genericlist[5];


            // If Excel is not Copied, Then Default Values are Set to the Total Assembly After Copy is Done.
            if (CopyMasterExcelOption == false)
            {
                String xlFilePath = System.IO.Path.ChangeExtension(m_assemblyFileName, ".xlsx");
                //if (xlFilePath == null || xlFilePath.Equals("") == true)
                //{
                //    e.Result = null;
                //    MessageBox.Show("Template is Missing in the Assembly");
                //    return;
                //}
                // 15-OCT change As per LTC Request, If XL File is Not Available, Then No Need to Read the EXCEL and Do Sync on DEFAULT VALUES.
                if (xlFilePath != null && xlFilePath.Equals("") == false && System.IO.File.Exists(xlFilePath) == true)
                {
                    ExcelData.readOccurenceVariablesFromTemplateExcel(xlFilePath, logFilePath);
                    //MessageBox.Show("Template is Missing in the Assembly");
                    //return;
                }
                
            }

            SolidEdgeData1.copyLinkedDocumentsToPublishedFolder2(DuplicateFolderPath, m_assemblyFileName, logFilePath, CopyMasterExcelOption);
            if (linkedDrawingsOption == true)
            {
                SolidEdgeData1.SearchAndcollectdrafts(m_assemblyFileName, DuplicateFolderPath, searchDrawingsFolder, logFilePath);
            }


            if (CopyMasterExcelOption == false)
            {
                String xlFilePath = System.IO.Path.ChangeExtension(m_assemblyFileName, ".xlsx");

                // 15-OCT change As per LTC Request, If XL File is Not Available, Then No Need to Read the EXCEL and Do Sync on DEFAULT VALUES.
                if (xlFilePath != null && xlFilePath.Equals("") == false && System.IO.File.Exists(xlFilePath) == true)
                {
                    try
                    {
                        String DupAssemblyFile = Path.Combine(DuplicateFolderPath, Path.GetFileName(m_assemblyFileName));
                        if (DupAssemblyFile == null || DupAssemblyFile.Equals("") == true)
                        {
                            e.Result = null;
                            return;
                        }
                        SolidEdgeInterface.SolidEdgeSync(DupAssemblyFile, logFilePath, "DEFAULTVALUE");
                    }
                    catch (Exception ex)
                    {

                        Utlity.Log("Exception: " + ex.Message, logFilePath);
                        Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
                    }
                }
            }

            try
            {
                // VariablePartsList - NULL, Suffix - ""
                ExcelComponentDeltaInterface.RenameComponentDetailsInMasterAssemblyTab(m_assemblyFileName, DuplicateFolderPath, null, "", logFilePath, "Duplicate");
            }
            catch (Exception ex)
            {
                Utlity.Log("RenameComponentDetailsInMasterAssemblyTab: " + ex.Message, logFilePath);
            }

            e.Result = genericlist;

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result !=null) 
            {
                List<object> genericlist = e.Result as List<object>;
                String DuplicateFolderPath = (String)genericlist[0];
                String m_assemblyFileName = (String)genericlist[1];
                String logFilePath = (String)genericlist[2];
                bool CopyMasterExcelOption = (bool)genericlist[3];
                bool linkedDrawingsOption = (bool)genericlist[4];
                String searchDrawingsFolder = (String)genericlist[5];

                Utlity.Log("-----------------------------------------------------------------", logFilePath);
                Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
                MessageBox.Show("Duplicate Documents To " + DuplicateFolderPath + " Done");
            }
            this.DialogResult = DialogResult.OK;

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Visible = true;
        }        

        

       
    }
}
