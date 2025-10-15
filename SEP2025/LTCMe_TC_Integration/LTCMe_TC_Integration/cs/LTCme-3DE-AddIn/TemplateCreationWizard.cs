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

namespace DemoAddInTC
{
    public partial class TemplateCreationWizard : Form
    {

        String PropertyFile = "";
        String ScanTemplateInFolder = "";
        String selectedTemplateFile = "";
        String home_Folder_Txt_File_Path = "";
        List<String> variablePartsList = new List<string>();
        int EXCEL_SHEET_NAME_LENGTH_LIMIT = 31; // Excel 2007

        // 28 - SEPT, Modified Method to getProperty
        public TemplateCreationWizard()
        {
            //MessageBox.Show(PropertyFile);
            PropertyFile = Utlity.getPropertyFilePath();
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                MessageBox.Show("Could Not Find Property File " + PropertyFile);
                return;
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                MessageBox.Show("Cannot Parse Property File to Get Templates Folder: " + PropertyFile);
                return;
            }
            ScanTemplateInFolder = Props.get("TEMPLATE_PUBLISH_FOLDER");
            home_Folder_Txt_File_Path = Props.get("CTD_TREEFILE");


            //if (ScanTemplateInFolder == null)
            //{
            //    MessageBox.Show("TEMPLATE_PUBLISH_FOLDER is not Set." + ScanTemplateInFolder);
            //    return;
            //}
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
                MessageBox.Show("InitializeComponent, Exception: " + ex.Message);
                return;
            }
            //try
            //{
            //    loadTemplatesToListView();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("InitializeComponent, Exception: " + ex.Message);
            //    return;
            //}

        }

        private void ApplySettings()
        {
            //this.templateDirectory.Text = ScanTemplateInFolder;
            //this.tabPage1.Hide();

            suffix.Text = "";
            suffix.Enabled = false;
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
                listView1.Items.Add(item);
                item.Tag = template;
                //i++;
            }


            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView1.Show();

        }



        private void TemplateCreationWizard_Load(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage2);
            //LoadTreeViewFromFile1(Ribbon2d.pathTotextFile, this.CTDTreeView);

            if (modeOfWorking.Equals("debug"))
                LoadTreeViewFromFile1("C:\\Users\\Administrator\\Downloads\\SEAssembly_IR.txt", this.CTDTreeView);
            else
            {
                LoadTreeViewFromFile1(home_Folder_Txt_File_Path, this.CTDTreeView);
                //LoadTreeViewFromFile1("C:\\ExchangeNew\\3DE\\SEAssemblyInfoFromHomeFolder\\SEAssembly_IR.txt", this.CTDTreeView);
            }


        }

        public static String FNDHOMEFOLDER = "fndoHomeFolder";
        public static String FNDOFOLDER = "fndoFolder";
        public static String ITEM_TYPE = "item";
        public static String ITEM_REVISION_TYPE = "item_revision";
        public static String NEW_STUFF_TYPE = "new_stuff";
        public static String TILDE = "~";

        private void LoadTreeViewFromFile1(string file_name, TreeView trv)
        {
            // Get the file's contents.
            string[] file_contents = File.ReadAllLines(file_name, Encoding.Default);

            // Process the lines.
            trv.Nodes.Clear();
            Dictionary<int, TreeNode> parents = new Dictionary<int, TreeNode>();
            foreach (string text_line in file_contents)
            {
                // Break the file into lines.
                string[] lines = text_line.Split(TILDE.ToCharArray());
                // See how many tabs are at the start of the line.


                int level = 0;
                foreach (string line in lines)
                {
                    String Name = line.Split('!')[0];
                    String Type = line.Split('!')[1];
                    String itemRevName = null;
                    int iType = 0;

                    if (Type.Equals(FNDHOMEFOLDER)) iType = 0;
                    if (Type.Equals(FNDOFOLDER)) iType = 1;
                    if (Type.Equals(ITEM_TYPE)) iType = 2;
                    if (Type.Equals(ITEM_REVISION_TYPE))
                    {
                        itemRevName = line.Split('!')[2];
                        Name = Name + "~" + itemRevName;
                        iType = 3;
                    }
                    if (Type.Equals(NEW_STUFF_TYPE)) iType = 4;

                    // Add the new node.
                    if (level == 0)
                    {
                        if (trv.Nodes.ContainsKey(Name) == false)
                        {
                            parents[level] = trv.Nodes.Add(Name, Name, iType, iType);
                        }
                    }
                    else
                    {
                        if (parents[level - 1].Nodes.ContainsKey(Name) == false)
                        {
                            parents[level] =
                                parents[level - 1].Nodes.Add(Name, Name, iType, iType);
                            parents[level].EnsureVisible();
                        }
                    }
                    level++;
                }
            }

            if (trv.Nodes.Count > 0) trv.Nodes[0].EnsureVisible();
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

        private string modeOfWorking = ""; //debug
        // Create Button
        private void button2_Click(object sender, EventArgs e)
        {
            if (CTDTreeView.SelectedNode != null)
            {
                String itemRevID = CTDTreeView.SelectedNode.Text.Split('~').First();
                String stageDir = Utlity.CreateLogDirectory();
                String logFilePath = System.IO.Path.Combine(stageDir, itemRevID.Replace("/", "_") + "_" + "CTD_1" + ".txt");

                Utlity.Log("---------------------------------\n", logFilePath, "CTD");
                Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString() + "\n", logFilePath, "CTD");
                List<object> arguments = new List<object>();
                arguments.Add(itemRevID);
                arguments.Add(logFilePath);
                arguments.Add(this.suffix.Text);
                arguments.Add(modeOfWorking);

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
            else
                MessageBox.Show("Select an item to clone");


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
            //if (this.selectedTemplateFile == "" || this.selectedTemplateFile == null || this.selectedTemplateFile.Equals("") == true)
            //{
            //    MessageBox.Show("Select The Template File in SELECT TEMPLATE Tab");
            //    return;
            //}
            if (this.CTDTreeView.SelectedNode == null || this.CTDTreeView.SelectedNode.Text.Equals("LTC Templates") == true)
            {
                MessageBox.Show("Select any Template File from the list");
                return;
            }
            else
            {
                TabPage t = tabControl1.TabPages[1];
                if (t != null)
                {
                    //tabControl1.SelectTab(t);
                    tabControl1.SelectedTab = t;
                }
                return;
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(tabControl1.SelectedTab.Text);
            if (tabControl1.SelectedTab.Text == "Derivative Creation Options")
            {
                if (this.CTDTreeView.SelectedNode == null || this.CTDTreeView.SelectedNode.Text.Equals("LTC Templates") == true)
                {
                    MessageBox.Show("Select any Template File from the list");

                    TabPage t = tabControl1.TabPages[0];
                    if (t != null)
                    {
                        //tabControl1.SelectTab(t);
                        tabControl1.SelectedTab = t;
                    }
                    return;
                }

                //loadVariablePartsToListView2();


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
            {
                ;
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
            extractor.GetThumbnail(LselectedTemplateFile, out hImageSE, false);
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

                string m_ImageFileLocation = Path.Combine(Path.GetTempPath(), "Image.jpg");
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
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true ||
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

        //private void backgroundWorker222_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    List<object> genericlist = e.Argument as List<object>;
        //    string ItemRevID = null;
        //    String logFilePath = (String)genericlist[1];
        //    String Suffix = (String)genericlist[2];
        //    string modeOfWorking = "";
        //    if (genericlist.Count == 4)
        //        modeOfWorking = (string)genericlist[3];

        //    Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath, "CTD");
        //    TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
        //    Utlity.Log("Initializing TC Services..", logFilePath, "CTD");
        //    TcAdaptor.TcAdaptor_Init(logFilePath);
        //    e.Result = "NOK";
        //    MessageBox.Show("Test Login Only..");

        //}

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            SolidEdgeFramework.Application objApp = null;
            SolidEdge.StructureEditor.Interop.Application StructureEditorApplication = null;
            SolidEdge.StructureEditor.Interop.SEECStructureEditor SEECStructure = null;
            SolidEdge.StructureEditor.Interop.ISEECStructureEditorATP atp = null;
            SolidEdge.StructureEditor.Interop.ISEECStructureEditor istr = null;
            SolidEdgeFramework.SolidEdgeTCE objSEEC = null;
            int commandResult = -1;
            String bstrFileName = "";
            String bstrRevisionRule = "Latest Working";
            String bstrFolderName = "";
            newItemFilePath = null;

            List<object> genericlist = e.Argument as List<object>;
            string ItemRevID = null;
            String logFilePath = (String)genericlist[1];
            String Suffix = (String)genericlist[2];
            string modeOfWorking = "";
            if (genericlist.Count == 4)
                modeOfWorking = (string)genericlist[3];

            try
            {
                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Checking if item and revision id is available from tree view selected node", logFilePath, "CTD")));
                ItemRevID = (String)genericlist[0];
                if (ItemRevID == null)
                    throw new Exception("Treeview Node text is null");
                if (ItemRevID.Split('/').Length != 2)
                    throw new Exception("Item and Revision ID could not be fetched from tree view");
                backgroundWorker1.ReportProgress(5);




                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("\nParsing selected node for item id and revision id", logFilePath, "CTD")));
                String bstrItemID = ItemRevID.Split('/').ElementAt(0);
                String bstrItemRevID = ItemRevID.Split('/').ElementAt(1);
                if (bstrItemID.Trim().Equals("") == true)
                    throw new Exception("Item ID from selected node is blank");
                if (bstrItemRevID.Trim().Equals("") == true)
                    throw new Exception("Revision ID from selected node is blank");
                backgroundWorker1.ReportProgress(10);


                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Acquire Existing solidedge session to download files", logFilePath, "CTD")));
                SolidEdgeFramework.Application currentSESession = SolidEdgeCommunity.SolidEdgeUtils.Connect(false);
                SolidEdgeFramework.SolidEdgeTCE currentSETCEObject = currentSESession.SolidEdgeTCE;
                objApp = currentSESession;

                //Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Starting new solidedge session and logging in to download files", logFilePath, "CTD")));
                //objApp = (SolidEdgeFramework.Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                if (objApp == null)
                    throw new Exception("Could not start new solid edge session");
                objSEEC = (SolidEdgeFramework.SolidEdgeTCE)objApp.SolidEdgeTCE;
                if (objSEEC == null)
                    throw new Exception("Could not get seec instance from solid edge session");
                objApp.DisplayAlerts = false;
                //objSEEC.SetTeamCenterMode(true);
                if (modeOfWorking.Equals("debug"))
                    objSEEC.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, "http://bednztcs05.soleras.local:8080/tc");
                else
                {
                    //objSEEC.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
                    //Utlity.Log("loginFromSE.group: " + loginFromSE.group, logFilePath);
                    //Utlity.Log("loginFromSE.role: " +  loginFromSE.role, logFilePath);
                    //Utlity.Log("loginFromSE.URL: " + loginFromSE.URL, logFilePath);

                    Utlity.Log("Login to SEEC is already DONE.", logFilePath);
                }
                backgroundWorker1.ReportProgress(20);




                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Getting bom of selected item " + bstrItemID, logFilePath, "CTD")));
                int NoOfComponents = 0;
                System.Object ListOfItemRevIds, ListOfFileSpecs;
                objSEEC.GetBomStructure(bstrItemID, bstrItemRevID, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                    out ListOfItemRevIds, out ListOfFileSpecs);
                Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                itemAndRevIds.Add(bstrItemID, bstrItemRevID);
                if (NoOfComponents > 0)
                {
                    System.Object[,] abcd = (System.Object[,])ListOfItemRevIds;
                    for (int i = 0; i < abcd.GetUpperBound(0); i++)
                    {
                        if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                            itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                    }
                }
                backgroundWorker1.ReportProgress(25);





                //Getting a new item id
                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Getting new item id for the top most assembly", logFilePath, "CTD")));
                string newItemID = null, newItemRevID = null;
                if (modeOfWorking.Equals("debug"))
                    objSEEC.AssignItemID("AI4_Item", out newItemID, out newItemRevID);
                else
                    objSEEC.AssignItemID("Ltc4_Item", out newItemID, out newItemRevID);
                if (newItemID == null || newItemRevID == null)
                    throw new Exception("Could not fetch new item id and rev id from teamcenter");
                Utlity.Log("newItemID " + newItemID, logFilePath);




                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Getting filename of assembly items", logFilePath, "CTD")));
                Dictionary<string, string> itemAndFileName = new Dictionary<string, string>();
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
                        if (filename.Contains(".par") || filename.Contains(".asm") || filename.Contains(".pwd") || filename.Contains(".psm"))
                        {
                            if (itemAndFileName.Keys.Contains(pair.Key) == false)
                                itemAndFileName.Add(pair.Key, filename);
                        }
                    }
                }
                backgroundWorker1.ReportProgress(50);





                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Starting structure editor", logFilePath, "CTD")));
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
                backgroundWorker1.ReportProgress(55);




                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Logging In to Teamcenter from structure editor", logFilePath, "CTD")));
                if (modeOfWorking.Equals("debug"))
                    commandResult = SEECStructure.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, "http://bednztcs05.soleras.local:8080/tc");
                else
                {
                    //commandResult = SEECStructure.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
                    //Utility.Log("loginFromSE.group: "+ loginFromSE.group, logFilePath);
                    
                    Utility.Log("ParsePropertyFile:" , logFilePath);
                    loginFromSE.ParsePropertyFile();
                    Utility.Log("loginFromSE.URL: " + loginFromSE.URL, logFilePath);
                    String URL = loginFromSE.URL;
                    Utility.Log("URL: " + URL, logFilePath);

                    string currentUserName = "";
                    string sepassword = "";
                    objSEEC.GetCurrentUserName(out currentUserName);
                    if (currentUserName == null || currentUserName.Equals("") == true)
                    {
                        Utility.Log("SEECStructure.ValidateLogin using dcproxy credentials..", logFilePath);
                        currentUserName = "dcproxy";
                        sepassword = "dcproxy";
                    } 
                    else
                    {
                        Utility.Log("SEECStructure.ValidateLogin using credentials.." + currentUserName, logFilePath);
                        sepassword = currentUserName;
                    }
                    commandResult = SEECStructure.ValidateLogin(currentUserName, sepassword, "Engineering", "Designer", URL);
                }

                if (commandResult != 0)
                    throw new Exception("Could not log in to teamcenter from structure editor");
                commandResult = SEECStructure.ClearCache();
                if (commandResult != 0)
                    throw new Exception("Could not clear structure editor cache");
                backgroundWorker1.ReportProgress(60);




                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Opening item to clone in structure editor", logFilePath, "CTD")));
                itemAndFileName.TryGetValue(bstrItemID, out bstrFileName);
                commandResult = SEECStructure.Open(bstrItemID, bstrItemRevID, bstrFileName, bstrRevisionRule, bstrFolderName);
                if (commandResult != 0)
                    throw new Exception("Cannot open " + bstrItemID + " in structure editor");
                backgroundWorker1.ReportProgress(65);




                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Settig action to save as all", logFilePath, "CTD")));
                commandResult = SEECStructure.SetSaveAsAll();
                
                if (commandResult != 0)
                    throw new Exception("Cannot assign action to save as in structure editor");
                backgroundWorker1.ReportProgress(70);



                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Setting item id for top most assembly", logFilePath, "CTD")));
                SEECStructure.SetDataIntoSingleCell(bstrItemID, bstrItemRevID, bstrFileName, "Item ID", newItemID);
                SEECStructure.SetDataIntoSingleCell(bstrItemID, bstrItemRevID, bstrFileName, "Revision", newItemRevID);



                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Setting new item id to all items", logFilePath, "CTD")));
                commandResult = SEECStructure.AssignAll();
                if (commandResult != 0)
                    throw new Exception("Cannot assign new item ids in structure editor");
                backgroundWorker1.ReportProgress(75);



                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Invoking run teamcenter validations", logFilePath, "CTD")));
                System.Object pvarListOfErrors, pvarListOfWarnings;
                commandResult = istr.RunTeamcenterValidation(out pvarListOfErrors, out pvarListOfWarnings);
                if (pvarListOfErrors != null)
                    throw new Exception("Run Teamcenter validations failed");
                backgroundWorker1.ReportProgress(80);



                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Performing actions and getting tal log file", logFilePath, "CTD")));
               // string bstrTALLogFileName = null;
                //atp.GetStructureEditorTALLogFileName(out bstrTALLogFileName);
                //MessageBox.Show(bstrTALLogFileName);
                commandResult = SEECStructure.PerformActions();
                if (commandResult != 0)
                    throw new Exception("Error encountered when performing actions");



                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Terminating structure editor", logFilePath, "CTD")));
                SEECStructure.Close();
                StructureEditorApplication.Quit();





                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Checking if new assembly " + newItemID + " is already in cache " + newItemID, logFilePath, "CTD")));
                Utlity.Log("Checking if new assembly " + newItemID + " is already in cache " + newItemID, logFilePath);
                string cachePath = null;
                objSEEC.GetPDMCachePath(out cachePath);
                Utlity.Log("cachePath " + cachePath, logFilePath);
                if (cachePath == null)
                    throw new Exception("Could not get cache Path");
                vFileNames = null;
                nFiles = 0;
                objSEEC.GetListOfFilesFromTeamcenterServer(newItemID, newItemRevID, out vFileNames, out nFiles);
                objArray = null;
                objArray = (System.Object[])vFileNames;
                bool topLevelFileExistsInCache = false;
                foreach (System.Object o in objArray)
                {
                    string filename = (string)(o);
                    Utlity.Log("filename " + filename, logFilePath);
                    if (filename.Contains(".par") || filename.Contains(".asm") || filename.Contains(".pwd") || filename.Contains(".psm"))
                    {
                        System.Object[,] temp = new object[1, 1];
                        if (File.Exists(Path.Combine(cachePath, filename)))
                        {
                            topLevelFileExistsInCache = true;
                            newItemFilePath = Path.Combine(cachePath, filename);
                            Utlity.Log("The new item's file path in cache is " + newItemFilePath, logFilePath, "CTD");
                            Utlity.Log("The new item's file path in cache is " + newItemFilePath, logFilePath);
                        }
                        else
                        {
                            topLevelFileExistsInCache = false;
                            newItemFilePath = Path.Combine(cachePath, filename);
                            Utlity.Log("The new item's file path in cache must be " + newItemFilePath, logFilePath, "CTD");
                            Utlity.Log("The new item's file path in cache must be " + newItemFilePath, logFilePath);
                        }
                    }
                }





                if (topLevelFileExistsInCache == false)
                {
                    Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Getting BOM structure of " + newItemID, logFilePath, "CTD")));
                    Utlity.Log("Getting BOM structure of " + newItemID, logFilePath, "CTD");
                    NoOfComponents = 0;
                    ListOfItemRevIds = null; ListOfFileSpecs = null;
                    objSEEC.GetBomStructure(newItemID, newItemRevID, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                        out ListOfItemRevIds, out ListOfFileSpecs);
                    itemAndRevIds = new Dictionary<string, string>();
                    itemAndRevIds.Add(newItemID, newItemRevID);
                    if (NoOfComponents > 0)
                    {
                        System.Object[,] tempObj = (System.Object[,])ListOfItemRevIds;
                        for (int i = 0; i < tempObj.GetUpperBound(0); i++)
                        {
                            if (itemAndRevIds.Keys.Contains(tempObj[i, 0].ToString()) == false)
                                itemAndRevIds.Add(tempObj[i, 0].ToString(), tempObj[i, 1].ToString());
                        }
                    }






                    Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Downloading all parts of " + newItemID, logFilePath, "CTD")));
                    Utlity.Log("Downloading all parts of " + newItemID, logFilePath, "CTD");
                    if (cachePath == null)
                        throw new Exception("Could not get cache path");
                    foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                    {
                        Utlity.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath, "CTD");
                        Utlity.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath);
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
                                Utlity.Log("Downloaded " + pair.Key + "/" + pair.Value, logFilePath, "CTD");
                                Utlity.Log("Downloaded " + pair.Key + "/" + pair.Value, logFilePath);
                            }
                        }
                    }
                }




                // 26-11-2024 | Murali | Comment this Logic
                //Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Terminating the instance of solidedge started for downloading files", logFilePath, "CTD")));
                //objApp.Quit();
                //objApp = null;





                SolidEdgeFramework.Application application = (SolidEdgeFramework.Application)SolidEdgeCommunity.SolidEdgeUtils.Connect();
                SolidEdgeFramework.Documents objDocuments = application.Documents;
                if (objDocuments == null)
                    throw new Exception("Object documents could not be loaded");
                SolidEdgeFramework.SolidEdgeTCE appSEEC = application.SolidEdgeTCE;
                string bStrCurrentUser = null;
                appSEEC.GetCurrentUserName(out bStrCurrentUser);
                Utlity.Log("bStrCurrentUser " + bStrCurrentUser, logFilePath);
                if (bStrCurrentUser.Equals(""))
                    appSEEC.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
                //Utlity.Log("Opening the new item in current solidedge session", logFilePath, "CTD");
                String PDMCachePathofExistingUser = "";
                appSEEC.GetPDMCachePath(out PDMCachePathofExistingUser);
                if (PDMCachePathofExistingUser == null)
                {
                    Utility.Log("Could not fetch Cache path. Trying again to login", "CTD");
                    Utility.Log("Could not fetch Cache path. Trying again to login", logFilePath);
                    appSEEC.ValidateLogin(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, loginFromSE.URL);
                    PDMCachePathofExistingUser = "";
                    appSEEC.GetPDMCachePath(out PDMCachePathofExistingUser);
                    if (PDMCachePathofExistingUser == null)
                    {
                        MessageBox.Show("Issue encountered when fetching cache path. Could not refresh login information. Please try again later in a new SE session");
                        e.Result = "NOK";
                        return;
                    }
                }
                else
                    Utility.Log("PDMCachePathofExistingUser " + PDMCachePathofExistingUser, logFilePath);

                if (!File.Exists(newItemFilePath))
                {
                    MessageBox.Show("New Item File does not exist..");
                    e.Result = "NOK";
                    return;
                }

                newItemFilePath = Path.Combine(PDMCachePathofExistingUser, Path.GetFileName(newItemFilePath));

                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Getting BOM structure of " + newItemID, logFilePath, "CTD")));
                NoOfComponents = 0;
                ListOfItemRevIds = null; ListOfFileSpecs = null;
                appSEEC.GetBomStructure(newItemID, newItemRevID, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                    out ListOfItemRevIds, out ListOfFileSpecs);
                itemAndRevIds = new Dictionary<string, string>();
                itemAndRevIds.Add(newItemID, newItemRevID);
                if (NoOfComponents > 0)
                {
                    System.Object[,] tempObj = (System.Object[,])ListOfItemRevIds;
                    for (int i = 0; i < tempObj.GetUpperBound(0); i++)
                    {
                        if (itemAndRevIds.Keys.Contains(tempObj[i, 0].ToString()) == false)
                            itemAndRevIds.Add(tempObj[i, 0].ToString(), tempObj[i, 1].ToString());
                    }
                }

                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Downloading items to existing User's Cache..", logFilePath, "CTD")));
                appSEEC.GetPDMCachePath(out cachePath);
                List<System.Object> FilesToCheckIn = new List<object>();
                Utlity.Log("Trying to download " + newItemID + "/" + newItemRevID, logFilePath, "CTD");
                vFileNames = null;
                nFiles = 0;
                appSEEC.GetListOfFilesFromTeamcenterServer(newItemID, newItemRevID, out vFileNames, out nFiles);
                objArray = null;
                objArray = (System.Object[])vFileNames;
                foreach (System.Object o in objArray)
                {
                    string filename = (string)(o);
                    if (filename.ToLower().Trim().Contains(".asm"))
                    {
                        System.Object[,] temp = new object[1, 1];
                        appSEEC.DownladDocumentsFromServerWithOptions(newItemID, newItemRevID, filename,
                            SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                            1, temp);
                        Utlity.Log("Downloaded " + newItemID + "/" + newItemRevID + " with all levels", logFilePath, "CTD");
                    }
                }

                /*foreach (KeyValuePair<string, string> pair in itemAndRevIds)
                {
                    Utlity.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath, "CTD");
                    vFileNames = null;
                    nFiles = 0;
                    appSEEC.GetListOfFilesFromTeamcenterServer(pair.Key, pair.Value, out vFileNames, out nFiles);
                    objArray = null;
                    objArray = (System.Object[])vFileNames;
                    foreach (System.Object o in objArray)
                    {
                        string filename = (string)(o);
                        if (filename.ToLower().Trim().Contains(".par") ||
                            filename.ToLower().Trim().Contains(".asm") ||
                            filename.ToLower().Trim().Contains(".pwd") ||
                            filename.ToLower().Trim().Contains(".psm") ||
                            filename.ToLower().Trim().Contains(".dft"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            appSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            FilesToCheckIn.Add(((System.Object)Path.Combine(cachePath, filename)));
                            Utlity.Log("Downloaded " + pair.Key + "/" + pair.Value, logFilePath, "CTD");
                        }
                    }
                }*/


                /*application.DisplayAlerts = false;
                Utlity.Log("Opening all the items and Checking in to Teamcenter", logFilePath, "CTD");
                foreach (System.Object o in FilesToCheckIn)
                {
                    string file = (string)o;
                    Utlity.Log("Opening " + file, logFilePath, "CTD");
                    SolidEdgeFramework.SolidEdgeDocument doc = (SolidEdgeFramework.SolidEdgeDocument)objDocuments.Open(file);
                    Utlity.Log("Opened " + file, logFilePath, "CTD");
                    System.Object[] objA = new object[1];
                    objA[0] = o;
                    appSEEC.CheckInDocumentsToTeamCenterServer(objA, false);
                    doc.Close();
                }*/

                application.DisplayAlerts = false;
                SolidEdgeAssembly.AssemblyDocument asyDoc = null;
                if (File.Exists(newItemFilePath) == true)
                {
                    Utlity.Log("Opening the new item in current solidedge session to check in the properties", logFilePath, "CTD");
                    objDocuments.Open(newItemFilePath);
                }
                else
                {
                    MessageBox.Show("File Missing in cache.." + newItemFilePath);
                    throw new Exception("File Missing in cache.." + newItemFilePath);
                }

                
                try
                {
                    Utlity.Log("Trying to get occurences", logFilePath, "CTD");
                    asyDoc = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument;
                    getOccurence(asyDoc, application);
                    Utlity.Log("Printing occurences", logFilePath, "CTD");
                    foreach (string s in uniqueOcc)
                    {
                        Utlity.Log(s, logFilePath, "CTD");
                    }
                    Utlity.Log("Converting to object", logFilePath, "CTD");
                    System.Object[] objA1 = new object[1];
                    objA1[0] = ((System.Object)asyDoc.FullName);
                    Utlity.Log("Checking in " + asyDoc.FullName, logFilePath, "CTD");
                    appSEEC.CheckInDocumentsToTeamCenterServer(objA1, false);
                    asyDoc.Close();

                    foreach (string s in uniqueOcc)
                    {
                        Utlity.Log("Opening " + s, logFilePath, "CTD");
                        SolidEdgeFramework.SolidEdgeDocument d = (SolidEdgeFramework.SolidEdgeDocument)objDocuments.Open(s);
                        Utlity.Log("Opened " + s, logFilePath, "CTD");
                        System.Object[] objA = new object[1];
                        objA[0] = ((System.Object)s);
                        appSEEC.CheckInDocumentsToTeamCenterServer(objA, false);
                        Utlity.Log("Checked in " + s, logFilePath, "CTD");
                        d.Close();
                        Utlity.Log("Closed " + s, logFilePath, "CTD");
                    }
                }
                catch (Exception ex)
                {
                    Utility.Log("Exception caught when trying to get occurences" + ex.ToString(), logFilePath);
                }
                application.DisplayAlerts = true;


                Utlity.Log("Opening the new item in current solidedge session", logFilePath, "CTD");
                objDocuments.Open(newItemFilePath);

                // parse the log file to check for any errors in it.
                String startString = "Utility Started @"; // to identify the start point of the log file
                bool parseLogFlag = Utility.parseLog(logFilePath, startString);

                Utlity.Log("Executing excel sanitize post clone", logFilePath, "CTD");
                Ribbon2d ribbon2dObj = new Ribbon2d();
                ribbon2dObj.RunSanitizeXL_PostClone_background();

                if (parseLogFlag == false)
                {
                    MessageBox.Show("1. Derivative Creation Completed but some errors were encountered during the process. Please check the log file " + logFilePath + "\n" +
                        "2. Item ID " + bstrItemID + " / " + bstrItemRevID + " cloned as " + newItemID + " / " + newItemRevID
                        + "\n" + "3. Please wait for Excel sanitization to complete." + "\n" +
                        "The progress will be shown in the prompt bar and a pop up will be shown after the process is complete. Please wait..");
                }
                else
                {
                    MessageBox.Show("Derivative Creation Completed. \nItem ID " + bstrItemID + "/" + bstrItemRevID + " cloned as " + newItemID + "/" + newItemRevID
                        + "\n" + "Please wait for Excel sanitization to complete." + "\n" +
                        "The progress will be shown in the prompt bar and a pop up will be shown after the process is complete. Please wait");
                    //MessageBox.Show("Derivative Creation Completed. \nItem ID " + bstrItemID + "//" + bstrItemRevID + " cloned as " + newItemID + "//" + newItemRevID
                    //   + "\n" + "Excel Sanitization completed");
                }
                backgroundWorker1.ReportProgress(100);
                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log("Utlity Ended @ " + System.DateTime.Now.ToString(), logFilePath, "CTD")));
                e.Result = "OK";

            }
            catch (Exception ex)
            {
                Ribbon2d.tcw.Invoke((Action)(() => Utlity.Log(ex.ToString(), logFilePath, "CTD")));
                try
                {
                    if (objApp != null)
                        objApp.Quit();
                }
                catch (Exception)
                {

                }

                try
                {
                    if (StructureEditorApplication != null)
                        StructureEditorApplication.Quit();
                }
                catch (Exception)
                {


                }


                e.Result = "NOK";
                if (ex.Message.Trim().Equals("") == false)
                    MessageBox.Show(ex.Message + "\nCTD could not be completed");
                else
                    MessageBox.Show(ex.ToString() + "CTD could not be completed");
            }

        }

        private List<string> uniqueOcc = new List<string>();
        private void getOccurence(SolidEdgeAssembly.AssemblyDocument asyDoc, SolidEdgeFramework.Application application )
        {
            foreach (SolidEdgeAssembly.Occurrence o in asyDoc.Occurrences)
            {
                string name = o.OccurrenceFileName;
                if (uniqueOcc.Contains(name) == false)
                {
                    uniqueOcc.Add(name);
                    if (name.ToLower().Trim().EndsWith(".asm"))
                    {
                        SolidEdgeAssembly.AssemblyDocument asyDoc1 = (SolidEdgeAssembly.AssemblyDocument)o.OccurrenceDocument;
                        getOccurence(asyDoc1, application);
                    }
                }
            }
        }

        public string newItemFilePath = null;
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result.Equals("OK") == true)
            {
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                this.DialogResult = DialogResult.OK;
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
            catch
            {
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

        private void statusTextBox_TextChanged(object sender, EventArgs e)
        {
            statusTextBox.Select(statusTextBox.TextLength, 1);
            //statusTextBox.SelectionStart = Ribbon2d.tcw.Text.Length;
            statusTextBox.ScrollToCaret();
            statusTextBox.Focus();
        }


    }
}
