using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DemoAddInTC.controller;
using DemoAddInTC.utils;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.se;

namespace DemoAddInTC.TC
{
    public partial class TC_TemplateCreationWizard : Form
    {
        
        String PropertyFile = "";
        String ScanTemplateInFolder = "";
        String selectedTemplateFile = "";
        List<String> variablePartsList = new List<string>();
        int EXCEL_SHEET_NAME_LENGTH_LIMIT = 31; // Excel 2007

        // 03-SEPT-2019, Added for Managed Mode.
        public TC_TemplateCreationWizard(String assemblyFileName)
        {


            this.selectedTemplateFile = assemblyFileName;

            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show("InitializeComponent, Exception: " + ex.Message);
                return;
            }

            try
            {
                ApplySettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ApplySettings, Exception: " + ex.Message);
                return;
            }
            
        }

        private void ApplySettings()
        {
            //this.templateDirectory.Text = ScanTemplateInFolder;
            //this.tabPage1.Hide();
            
            suffix.Text = "";
            suffix.Enabled = false;

            TabPage t = tabControl1.TabPages[1];
            if (t != null)
            {
                tabControl1.SelectTab(t);
            }
        }



        private void loadTemplatesToListView()
        {
            ScanTemplateInFolder = this.templateDirectory.Text;

            List<String> templateFiles = scanTemplatesInFolder(ScanTemplateInFolder);
            if (templateFiles == null || templateFiles.Count == 0)
            {
                MessageBox.Show("NO TEMPLATE FILES FOUND");
                return;
            }
            //listView1.SmallImageList = imageList1;
            //listView1.Columns.Add("No", -2, HorizontalAlignment.Left);
            listView1.Columns.Add("Template Name", -2, HorizontalAlignment.Left);
            listView1.Columns.Add("Template Path", -2, HorizontalAlignment.Left);
            listView1.View = View.Details;
            //int i = 1;
            //set the Small and large ImageList properties of listview
            //listView1.LargeImageList = imageList1;
            listView1.SmallImageList = imageList1;
            foreach (String template in templateFiles)
            {
                //String templateAssemFile = System.IO.Path.ChangeExtension(template, ".asm");
                String file = System.IO.Path.GetFileName(template);
                String directoryName = System.IO.Path.GetDirectoryName(template);                
                String[] fileArr = new string[2];
                //fileArr[0] = i.ToString();
                fileArr[0] = file;
                fileArr[1] = directoryName;
                ListViewItem item = new ListViewItem(fileArr);
                item.ImageIndex = 0;
                listView1.Items.Add(item );
                item.Tag = template;
                //i++;
            }
            

            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView1.Show();
            
        }

       

        private void TC_TemplateCreationWizard_Load(object sender, EventArgs e)
        {

        }

        //Close
        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        //Close
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Back In TabPage2
        private void button1_Click(object sender, EventArgs e)
        {
            TabPage t = tabControl1.TabPages[0];
            if (t != null)
            {
                tabControl1.SelectTab(t);
            }
        }

        // Create Button
        private void button2_Click(object sender, EventArgs e)
        {
            
            String folderToPublish = this.publishFolder.Text;
            if (folderToPublish == null || folderToPublish.Equals("") == true)
            {
                MessageBox.Show("Choose the Folder to Create Derivative");
                return;
            }

            if (Directory.Exists(folderToPublish) == false)
            {
                MessageBox.Show("Directory does not Exist.." + folderToPublish + " .. Please Browse & Set the Path Again..");
                return;
            }
            string[] filePaths = Directory.GetFiles(folderToPublish, "*.*",
                                         SearchOption.AllDirectories);

            if (filePaths != null && filePaths.Length > 0)
            {
                MessageBox.Show("Clear the Derivative Directory & Proceed to do Create.");
                return;
            }

            String assemblyFileName  = this.selectedTemplateFile;
            if (assemblyFileName == null || assemblyFileName.Equals("") == true)
            {
                MessageBox.Show("Choose the Template Assembly File to Create Derivative");
                return;
            }
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "CTD_1" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            // 29 - SEPT
            variablePartsList = RemoveVariablePartsIfNameExceedsLimit(variablePartsList, suffix.Text);
            List<object> arguments = new List<object>();
            arguments.Add(folderToPublish);
            arguments.Add(assemblyFileName);
            arguments.Add(variablePartsList);
            arguments.Add(logFilePath);
            arguments.Add(this.suffix.Text);

            if (backgroundWorker1.IsBusy != true)
            {
                this.groupBox5.Visible = true;
                this.progressBar1.Visible = true;
                ViewUtils.EnableDisableAllControls(false, this);
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {

                Utlity.Log("Background Worker Thread is Busy, CTD", logFilePath);
            }

            
        }

        // 29 SEPT -- If PartName + Suffix Exceeds 31, Its Removed from the List
        private List<string> RemoveVariablePartsIfNameExceedsLimit(List<string> variablePartsList, string Suffix)
        {
            variablePartsList.RemoveAll(PartName => (PartName + Suffix).Length > EXCEL_SHEET_NAME_LENGTH_LIMIT);
            return variablePartsList;
        }
        

        // Next - In TabPage1
        private void button5_Click(object sender, EventArgs e)
        {
            if (this.selectedTemplateFile == "" || this.selectedTemplateFile == null || this.selectedTemplateFile.Equals("") == true)
            {
                MessageBox.Show("Select The Template File in SELECT TEMPLATE Tab");
                return;
            }
            else
            {
                TabPage t = tabControl1.TabPages[1];
                if (t != null)
                {
                    tabControl1.SelectTab(t);
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(tabControl1.SelectedTab.Text);
            if (tabControl1.SelectedTab.Text == "Select Template")
            {
               
                {
                    MessageBox.Show("SELECT TEMPLATE Option is Disabled");

                    TabPage t = tabControl1.TabPages[1];
                    if (t != null)
                    {
                        tabControl1.SelectTab(t);
                    }
                    return;
                }

                


            }

            if (tabControl1.SelectedTab.Text == "Derivative Creation Options")
            {
                if (this.selectedTemplateFile == "" || this.selectedTemplateFile == null || this.selectedTemplateFile.Equals("") == true)
                {
                    MessageBox.Show("Select The Template File in SELECT TEMPLATE Tab");

                    TabPage t = tabControl1.TabPages[0];
                    if (t != null)
                    {
                        tabControl1.SelectTab(t);
                    }
                    return;
                }

                loadVariablePartsToListView2();
           
                
            }

        }

        private void loadVariablePartsToListView2()
        {
            String assemblyFileName = this.selectedTemplateFile;

            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "CTD_2" + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);            
            variablePartsList = readVariablePartsFromTemplate(assemblyFileName, logFilePath);            
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);

            listView2.Clear();
            listView2.Columns.Add("No", -2, HorizontalAlignment.Left);
            listView2.Columns.Add("Variable Part Name", -2, HorizontalAlignment.Left);

            listView2.View = View.Details;
            int i = 1;
            foreach (String variablePart in variablePartsList)
            {                
                String[] fileArr = new string[3];
                fileArr[0] = i.ToString();
                fileArr[1] = variablePart;                
                ListViewItem item = new ListViewItem(fileArr);
                listView2.Items.Add(item);
                item.Tag = variablePart;
                i++;
            }
            listView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView2.Show();

        }

        private List<string> GetFullPathforVariablePartsInTemplateDirectory(List<string> variablePartNamesList)
        {
            throw new NotImplementedException();
        }

        

        // Browse - TabPage1 - Select Template
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {                ;
                this.templateDirectory.Text = folderBrowserDialog1.SelectedPath;
            }

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listView1.SelectedItems.Count == 0)
            {
                selectedTemplateFile = "";
                return;
            }

            selectedTemplateFile = this.listView1.SelectedItems[0].Tag.ToString();
            //MessageBox.Show(selectedTemplateFile);
            List<object> arguments = new List<object>();
            arguments.Add(selectedTemplateFile);

            if (backgroundWorker2.IsBusy != true)
            {
                backgroundWorker2.RunWorkerAsync(arguments);
            }
            else
            {
                MessageBox.Show("Busy, Try Again");
                return;
            }
            
        }

        private void getViewForTemplateFile1(String LselectedTemplateFile)
        {
            var extractor = new SeThumbnailLib.SeThumbnailExtractor();
            int hImageSE = 0;
            extractor.GetThumbnail(LselectedTemplateFile, out hImageSE,false);
            Image image = Image.FromHbitmap(new IntPtr(hImageSE));
            ClearPicture();
            SetPicture(image);

        }

        private void getViewForTemplateFile(String LselectedTemplateFile)
        {
            ClearPicture();

            if (LselectedTemplateFile == null || LselectedTemplateFile.Equals("") == true)
            {
                return;
            }
            
            SolidEdgeFramework.View view = null;
            
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
                return;

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;
            objApp.DisplayAlerts = false;            
            //objApp.Visible = false;
            objDocuments = objApp.Documents;
            if (objDocuments == null)
            {
                Utlity.ResetAlerts(objApp, true, "");
                return;
            }
            
            objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(LselectedTemplateFile);
            if (objAssemblyDocument == null)
            {
                Utlity.ResetAlerts(objApp, true, "");
                return;
            }
            SolidEdgeFramework.Window window = (SolidEdgeFramework.Window)objApp.ActiveWindow;
            if (window != null)
            {
                view = window.View;
                
                string m_ImageFileLocation = Path.Combine(Path.GetTempPath(),"Image.jpg");
                object Width = 1920;
                object Height = 1080;
                object AltViewStyle = "Default";
                object Resolution = 1;
                object ColorDepth = 24;
                var ImageQuality = SolidEdgeFramework.SeImageQualityType.seImageQualityHigh;
                bool Invert = false;
                if (File.Exists(m_ImageFileLocation) == true)
                    File.Delete(m_ImageFileLocation);
                view.SaveAsImage(m_ImageFileLocation, Width, Height, AltViewStyle, Resolution, ColorDepth, ImageQuality, Invert);

                if (System.IO.File.Exists(m_ImageFileLocation) == true)
                {
                    Image tmp = Image.FromFile(m_ImageFileLocation);
                    SetPicture(tmp);
                }
            }
            objAssemblyDocument.Close();
            objAssemblyDocument = null;
            Utlity.ResetAlerts(objApp, true, "");
            window.Close();
            window = null;

            

        }

        private void ClearPicture()
        {
            if (pictureBox1.InvokeRequired)
            {
                pictureBox1.Invoke(new MethodInvoker(
                delegate()
                {
                    if (pictureBox1.Image != null)
                    {
                        pictureBox1.Image.Dispose();
                        pictureBox1.Invalidate();
                        pictureBox1.Image = null;
                        pictureBox1.InitialImage = null;
                    }
                }));
            }
            else
            {
                if (pictureBox1.Image != null)
                {
                    pictureBox1.Image.Dispose();
                    pictureBox1.Invalidate();
                    pictureBox1.Image = null;
                    pictureBox1.InitialImage = null;
                }
            }
        }

        private void SetPicture(Image img)
        {
            if (pictureBox1.InvokeRequired)
            {
                pictureBox1.Invoke(new MethodInvoker(
                delegate()
                {
                    pictureBox1.Image = img;
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                }));
            }
            else
            {
                pictureBox1.Image = img;
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {                
                this.publishFolder.Text = folderBrowserDialog1.SelectedPath;
            }

        }

        private List<String> readVariablePartsFromTemplate(string AssemblyFilePath, String logFilePath)
        {
            List<String> variablePartsList = new List<String>();
            Utlity.Log("----------------------------------------------------------", logFilePath);
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            String xlFilePath = Path.ChangeExtension(AssemblyFilePath, ".xlsx");

            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            if (f.Exists == true)
            {
                workbooks = xlApp.Workbooks;
                Utlity.Log("File Already Exists" + xlFilePath, logFilePath);
                //xlApp.DisplayAlerts = true;
                try
                {
                    xlWorkbook = workbooks.Open(xlFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = workbooks.Open(xlFilePath, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        MessageBox.Show(ex1.Message);
                    }
                    
                }
                xlApp.DisplayAlerts = false;
            }
            else
            {
                Utlity.Log("file does not Exist: " + xlFilePath, logFilePath);
                return null;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return null;
            }
            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);


            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        Utlity.Log("Skipping Custom LTC Sheet: " + sheet.Name, logFilePath);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }

                Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals("MASTER ASSEMBLY",StringComparison.OrdinalIgnoreCase) == true  ||
                    sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet); 
                    continue;
                }
                if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
                {
                    variablePartsList.Add(sheet.Name);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);        
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background            
            //Marshal.ReleaseComObject(xlWorksheet);

            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Close(true);


            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
            sheets = null;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            workbooks = null;

            Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlApp != null) xlApp.DisplayAlerts = true;
            //quit and release
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            return variablePartsList;
            
        }

        private List<String> scanTemplatesInFolder(string ScanTemplateInFolder)
        {
            List<String> AssemblyTemplateFilesList = new List<string>();
            String[] excelFiles = Directory.GetFiles(ScanTemplateInFolder, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".xlsx")))
                                         .ToArray();
            if (excelFiles == null || excelFiles.Length == 0)
            {
                //MessageBox.Show("NO TEMPLATES FOUND");
                return null;
            }
            foreach (String excelFile in excelFiles)
            {
                String AssemblyFileWithTemplate = Path.ChangeExtension(excelFile, ".asm");
                if (File.Exists(AssemblyFileWithTemplate) == true)
                {
                    if (AssemblyTemplateFilesList.Contains(AssemblyFileWithTemplate) == false)
                    {
                        AssemblyTemplateFilesList.Add(AssemblyFileWithTemplate);
                    }
                }

            }

            return AssemblyTemplateFilesList;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
            List<object> genericlist = e.Argument as List<object>;
            String folderToPublish = (String)genericlist[0];
            String assemblyFileName = (String)genericlist[1];
            List<String> variablePartsList = (List<String>)genericlist[2];
            String logFilePath = (String)genericlist[3];
            String Suffix = (String)genericlist[4];
            backgroundWorker1.ReportProgress(25);
            String bstrItemId = "";
            String bstrItemRev = "";
            SEECAdaptor.LoginToTeamcenter(logFilePath);
            SolidEdgeFramework.SolidEdgeTCE SolidEdgeTCE=   SEECAdaptor.getSEECObject();
            SolidEdgeTCE.GetDocumentUID(assemblyFileName, out bstrItemId, out bstrItemRev);
            SEECStuctureEditor.perform_clone(bstrItemId,bstrItemRev,logFilePath);
            //SolidEdgeData1.copyLinkedDocumentsToPublishedFolder2(folderToPublish, assemblyFileName, logFilePath, true);
            backgroundWorker1.ReportProgress(50);
            //try
            //{
            //    SolidEgeData2.CopyAndReplaceSuffixForVariableParts3(folderToPublish, assemblyFileName, variablePartsList, Suffix, logFilePath);
            //}
            //catch (Exception ex)
            //{
            //    Utlity.Log("CopyAndReplaceSuffixForVariableParts3: " + ex.Message, logFilePath);
            //}

            
            backgroundWorker1.ReportProgress(75);

            //try
            //{
            //    ExcelComponentDeltaInterface.RenameComponentDetailsInMasterAssemblyTab(assemblyFileName, folderToPublish, variablePartsList, Suffix, logFilePath, "CTD");
            //}
            //catch (Exception ex)
            //{
            //    Utlity.Log("RenameComponentDetailsInMasterAssemblyTab: " + ex.Message, logFilePath);
            //}

            // 28-OCT
            //try
            //{
            //    ExcelDeltaInterface.RenameVariablePartsInFeaturesTab(assemblyFileName, folderToPublish, variablePartsList, Suffix, logFilePath);
            //}
            //catch (Exception ex)
            //{
            //    Utlity.Log("RenameVariablePartsInFeaturesTab: " + ex.Message, logFilePath);
            //}


            // 21- SEPT - 2018
            // Update The Suffix to Excel SheetName & partName Column in Each Sheet.
            // Issue -1 in Designer Procedure, After CTD is done, SyncTE and Working on Enabling/Disabling Features Wont Work.
            // This is due to the Sheet/PartNames not uptoDate with the Suffix that is changed in CTD.
            //try
            //{
            //    ExcelInterface.RenameSheetNamesForVariableParts(folderToPublish,assemblyFileName, variablePartsList, Suffix, logFilePath);
            //}
            //catch (Exception ex)
            //{
            //    Utlity.Log("RenameSheetNamesForVariableParts: " + ex.Message, logFilePath);
            //}

            


            // 15-OCT -- Search and Copy Drafts to the Custom Template Create Folder
            Utlity.Log("Copying Associated Drafts to Publish Folder: " + folderToPublish, logFilePath);
            String searchDrawingsFolder = System.IO.Path.GetDirectoryName(assemblyFileName);
            Utlity.Log("searchDrawingsFolder: " + searchDrawingsFolder, logFilePath);
            //SolidEdgeData1.SearchAndcollectdrafts(assemblyFileName, folderToPublish, searchDrawingsFolder, logFilePath);
            //SolidEdgeRedefineLinks.ReplaceLinks1(folderToPublish,variablePartsList, Suffix, logFilePath);
            Utlity.Log("Custom Template Creation Completed " + System.DateTime.Now.ToString(), logFilePath);

            backgroundWorker1.ReportProgress(100);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            e.Result = "OK";

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result.Equals("OK") == true)
            {
                this.DialogResult = DialogResult.OK;
                MessageBox.Show("Derivative Creation Completed");
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Visible = true;
            this.progressBar1.Value = e.ProgressPercentage;

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                suffix.Enabled = true;
            }
            else
            {
                suffix.Text = "";
                suffix.Enabled = false;
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String selectedTemplateFile = (String)genericlist[0];
            try
            {
                getViewForTemplateFile1(selectedTemplateFile);
            }
            catch {
            }
            e.Result = "OK";
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result.Equals("OK") == true)
            {
               // MessageBox.Show("CTD Completed");
            }

        }

       
    }
}
