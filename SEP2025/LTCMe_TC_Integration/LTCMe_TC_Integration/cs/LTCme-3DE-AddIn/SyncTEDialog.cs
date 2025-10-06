using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.utils;
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
using DemoAddInTC.se;
using Microsoft.Office.Interop.Excel;

namespace DemoAddInTC
{
    // 28 Sept, IMPORTANT - Trace Listener Would Work Only in DEBUG mode.
    // Build the DLL in DLL Mode to Take Advantage of this.
    public partial class SyncTEDialog : Form
    {
        public SyncTEDialog()
        {
            InitializeComponent();

            Trace.Listeners.Add(new ListBoxTraceListener(listBox1));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.progressBar1.Visible = true;
            SyncVariablesFromExcelToSolidEdge();

        }

        public static List<string> listOfFileNamesInSession = new List<string>();
        public static SolidEdgeFramework.Application application = null;
        public static SolidEdgeDocument document = null;
        // Sync the Variables From Excel To SolidEdge
        private void SyncVariablesFromExcelToSolidEdge()
        {
            listOfFileNamesInSession = new List<string>();


            // Connect to running Solid Edge Instance
            application = SE_SESSION.getSolidEdgeSession();
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

            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            String AssemstageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + "SyncTE" + ".txt");








            String xlFile = System.IO.Path.Combine(AssemstageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            Utlity.Log("xlFile: " + xlFile, logFilePath);
            Utlity.Log("logFilePath: " + logFilePath, logFilePath);
            Utlity.Log("assemblyFileName: " + assemblyFileName, logFilePath);
            List<object> arguments = new List<object>();
            arguments.Add(xlFile);
            arguments.Add(logFilePath);
            arguments.Add(assemblyFileName);

            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker is Busy, SyncTE", logFilePath);
            }

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String xlFile = (String)genericlist[0];
            String logFilePath = (String)genericlist[1];
            String assemblyFileName = (String)genericlist[2];
            //if (System.IO.File.Exists(xlFile) == false)
            //{
            //    this.progressBar1.Visible = false;
            //    Utlity.Log("File does not Exist: " + xlFile, logFilePath);
            //    e.Result = null;
            //    return;
            //}

            if (System.IO.File.Exists(assemblyFileName) == false)
            {
                this.progressBar1.Visible = false;
                Utlity.Log("File does not Exist: " + assemblyFileName, logFilePath);
                e.Result = null;
                return;
            }

            if (System.IO.File.Exists(logFilePath) == false)
            {
                this.progressBar1.Visible = false;
                Utlity.Log("File does not Exist: " + logFilePath, logFilePath);
                e.Result = null;
                return;
            }



            //Finding documents of current assembly
            try
            {
                SolidEdgeFramework.SolidEdgeTCE objSEEC = application.SolidEdgeTCE;
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
                        if (filename.Contains(".dft"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", true, false,
                                1, temp);
                            Utlity.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }
                    }
                }

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
                        if (filename.Contains(".par") || filename.Contains(".asm") || filename.Contains(".pwd") || filename.Contains(".psm") || filename.Contains(".dft"))
                        {
                            if (listOfFileNamesInSession.Contains(filename) == false)
                                listOfFileNamesInSession.Add(filename.ToLower().Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.Log("An exception was caught while trying to fetch list of file names " + ex.ToString(), logFilePath);

            }
            if (listOfFileNamesInSession.Count != 0)
            {
                Utility.Log("Printing listOfFileNamesInSession", logFilePath);
                foreach (string s in listOfFileNamesInSession)
                    Utility.Log(s, logFilePath);
            }
            else
                Utility.Log("No documents found in listOfFileNamesInSession", logFilePath);



            //Opening excel application for remove component purpose
            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.DisplayAlerts = false;
                if (File.Exists(xlFile) == false)
                {
                    Utlity.Log("download Excel Template From Teamcenter..", logFilePath, "INFO");
                    Ribbon2d.downloadExcelTemplateFromTeamcenter(assemblyFileName, logFilePath);
                }

                if (File.Exists(xlFile))
                {
                    workbooks = xlApp.Workbooks;
                    Utlity.Log("File Already Exists", logFilePath);
                    xlApp.DisplayAlerts = false;
                    //xlWorkbook = workbooks.Open(xlFilePath);
                    try
                    {
                        xlWorkbook = workbooks.Open(xlFile);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            xlWorkbook = workbooks.Open(xlFile, CorruptLoad: 1);
                        }
                        catch (Exception ex1)
                        {
                            System.Windows.Forms.MessageBox.Show(ex1.Message);
                        }

                    }
                }
                else
                {
                    Utlity.Log("file does not Exist: " + xlFile, logFilePath);
                    e.Result = null;
                    return;
                }
                if (xlWorkbook == null)
                {
                    Utlity.Log("xlWorkBook is NULL", logFilePath);
                    return;
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Exception " + ex.ToString(), logFilePath);
                e.Result = "NOK";
                return;
            }


            //Reading list of components in master sheet 
            try
            {
                Utlity.Log("ReadMasterAssemblySheet: ", logFilePath);
                ReadMasterAssemblySheet(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("ReadMasterAssemblySheet: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }

            //Removing excluded components sheets in excel
            try
            {
                Utlity.Log("RemoveSheet: ", logFilePath);
                RemoveSheet(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("RemoveSheet: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }

            //Removing excluded components in features sheet
            try
            {
                Utlity.Log("RemoveComponentsFromFeatureTab: ", logFilePath);
                RemoveComponentsInFeatureTAB(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("RemoveComponentsFromFeatureTab: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }

            try
            {
                Utlity.Log("Saving the Changes Done..", logFilePath);
                Utlity.Log("Saving the Changes Done..", logFilePath, "INFO");
                xlApp.ActiveWorkbook.Save();
            }
            catch (Exception)
            {
                Utlity.Log("Cannot save the changes in excel", logFilePath);
            }

            //Removing excluded components from Solidedge
            try
            {
                //Getting solidedge session
                SolidEdgeFramework.Application Seapplication = null;
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


                Utlity.Log("Deleting Occurrences In SolidEdge", logFilePath, "INFO");
                Utlity.Log("Deleting Occurrences In SolidEdge", logFilePath);
                SolidEdgeOccurenceDelete_1 occDelete = new SolidEdgeOccurenceDelete_1();
                occDelete.SolidEdgeOccurrenceDeleteFromExcelSTAT(topLineAssembly, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SolidEdgeOccurrenceDeleteFromExcelSTAT: " + ex.Message, logFilePath);
                e.Result = "NOK";
                return;
            }

            //Closing the excel application that was opened
            try
            {
                xlWorkbook.Close(true);

                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;

                Marshal.ReleaseComObject(workbooks);
                workbooks = null;

                Utlity.Log("Workbooks opned for remove component are closed ", logFilePath);
                xlApp.DisplayAlerts = true;
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
            catch (Exception)
            {
                Utility.Log("Cannot close the excel application started for remove component", logFilePath);
            }



            
            try
            {
                Utlity.Log("Reading Data from Excel..", logFilePath, "INFO");
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
                Utlity.Log("Syncing Data To Solid Edge..", logFilePath, "INFO");
                SolidEdgeInterface.SolidEdgeSync(assemblyFileName, logFilePath, "VALUE");
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
            }

            try
            {
                Utlity.Log("Syncing Features to Solid Edge", logFilePath, "INFO");
                SyncFeaturesToSolidEdge(xlFile, assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                e.Result = "NOK";
                return;
            }

            // Commented On 4 Dec 2018 - On Request of Simone.
            //try
            //{
            //    Utlity.Log("ReConnect To Solid Edge..", logFilePath, "INFO");
            //    ConnectToSolidEdge(logFilePath);
            //}
            //catch (Exception ex)
            //{
            //    e.Result = "NOK";
            //    return;
            //}

            //try
            //{
            //    ModifyXL(xlFile, logFilePath);
            //}
            //catch (Exception ex)
            //{
            //    e.Result = "NOK";
            //    return;
            //}

            // Update the Drafts
            try
            {
                Utlity.Log("Updating Views in Draft Files....", logFilePath, "INFO");
                SolidEdgeUpdateView.SearchDraftFile(assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                e.Result = "NOK";
                return;
            }

            String tc_mode = Utlity.getManageMode(logFilePath);
            if (tc_mode.Equals("YES", StringComparison.OrdinalIgnoreCase) == true)
            {
                // Upload the files back to Teamcenter - 17 August 2019
                try
                {
                    
                    Utlity.Log("SEEC Login....", logFilePath, "INFO");
                    SEECAdaptor.LoginToTeamcenter(logFilePath);

                    string bStrCurrentUser = null;
                    SEECAdaptor.getSEECObject().GetCurrentUserName(out bStrCurrentUser);

                    String password = bStrCurrentUser;

                    Utlity.Log("Uploading files back to Teamcenter....", logFilePath, "INFO");
                    Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath);
                    Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath);
                    Utlity.Log("ID=" + bStrCurrentUser, logFilePath);
                    Utlity.Log("Group=Engineering", logFilePath);
                    Utlity.Log("Role=Designer", logFilePath);
                    TcAdaptor.login(bStrCurrentUser, password, "Engineering", "Designer", logFilePath);
                    Utlity.Log("Initializing TC Services..", logFilePath);
                    //bool logIn_Success = TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
                    TcAdaptor.TcAdaptor_Init(logFilePath);
                    Utlity.Log("Upload Excel to Teamcenter....", logFilePath, "INFO");
                    TcAdaptor.uploadExcelToTC(bStrCurrentUser, password, "Engineering", "Designer", xlFile, logFilePath);
                    Utlity.Log("checking in SE documents to Teamcenter....", logFilePath, "INFO");
                    SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath);
                    Utlity.Log("logout from TC", logFilePath, "INFO");
                    TcAdaptor.logout(logFilePath);
                }
                catch (Exception ex)
                {
                    e.Result = "NOK";
                    return;
                }
            }

            // Murali - 25-NOV-2024 - SOA Decustomization - Start
            //if (tc_mode.Equals("YES", StringComparison.OrdinalIgnoreCase) == true)
            //{
            //    // Upload the files back to Teamcenter using SEEC API and not using SOA - 25-NOV-2024
            //    try
            //    {
            //        Utlity.Log("Uploading files back to Teamcenter....", logFilePath, "INFO");
            //        Utlity.Log("SEEC Login....", logFilePath, "INFO");
            //        SEECAdaptor.LoginToTeamcenter(logFilePath);
            //        Utlity.Log("checking in SE documents to Teamcenter....", logFilePath, "INFO");
            //        SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath);
            //        Utlity.Log("Upload Excel to Teamcenter using SEEC....", logFilePath, "INFO");
            //        Ribbon2d.UploadtoTCUsingSEEC(Ribbon2d.currentSETCEObject, Ribbon2d.currentSESession);
            //    }
            //    catch (Exception ex)
            //    {
            //        e.Result = "NOK";
            //        return;
            //    }
            //}
            // Murali - 25-NOV-2024 - SOA Decustomization - End

            Utlity.Log("SyncTE completed..", logFilePath, "INFO");
            e.Result = genericlist;

        }


        // Read the Master Assembly to Collect the Components that Needs to be Added/Removed
        public static List<String> componentList = new List<String>();
        // 09-10-2024 | Murali || START
        // Added Logic for Request from LTC (ALLEN). Sync 3D must ignore, not remove and do nothing on parts/assemblies that are not present in Excel Template “MASTER ASSEMBLY” sheet
        // 09-10-2024 | Murali || END
        public static List<String> ExcludecomponentList = new List<String>();
        public static void ReadMasterAssemblySheet(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            componentList.Clear();
            ExcludecomponentList.Clear();
            //Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utlity.Log("xlApp is NULL", logFilePath);
                return;

            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }

            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                //Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                    //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i == 1)
                            continue;

                        String Status = "";

                        //6 is the Status
                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {


                            try
                            {
                                Status = xlRange.Cells[i, 6].Value2;
                                //Utlity.Log(Status, logFilePath);

                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Status" + ex.Message, logFilePath);
                            }


                        }
                        // FullName (Includes Path)
                        if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // FullName
                                filePath = xlRange.Cells[i, 7].Value2;
                                String fileName = Path.GetFileName(filePath);
                                //Utlity.Log(fileName, logFilePath);
                                if (componentList.Contains(fileName) == false)
                                {
                                    if (Status != null && Status.Equals("") == false)
                                    {
                                        if (Status.Equals("INCLUDED", StringComparison.OrdinalIgnoreCase) == true)
                                        {
                                            componentList.Add(fileName);
                                        }
                                        else if (Status.Equals("EXCLUDED", StringComparison.OrdinalIgnoreCase) == true)
                                        {
                                            Utlity.Log("fileName is Excluded : " + fileName, logFilePath);
                                            if (ExcludecomponentList.Contains(fileName) == false)
                                            {
                                                ExcludecomponentList.Add(fileName);
                                            }
                                        }

                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("FullName" + ex.Message, logFilePath);
                            }
                        }

                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);

                }
                else
                {
                    continue;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);

            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //xlApp.Visible = true;     

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;
            //Utlity.Log("----------------------------------------------------------", logFilePath);
            // Release xlApp outside this Function.
        }




        public static void RemoveComponentsInFeatureTAB(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            List<String> MasterAssemblyList = componentList;

            if (MasterAssemblyList == null || MasterAssemblyList.Count == 0)
            {
                Utlity.Log("RemoveComponentsInFeatureTAB: " + "Remove Component List is Empty", logFilePath);
                return;
            }

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utlity.Log("xlApp is NULL", logFilePath);
                return;

            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            Utlity.Log("Removing Variable Parts In Feature Tab, If User does Not Need It...", logFilePath);

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {

                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                    //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = xlRange.Rows.Count; i > 1; i--)
                    {
                        if (i == 1)
                            continue;

                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {
                            String componentName = xlRange.Cells[i, 6].Value2;

                            if (MasterAssemblyList.Contains(componentName) == false)
                            {
                                //Utlity.Log("componentName: " + componentName, logFilePath);
                                Range range = sheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                                //setting the range for deleting the rows
                                range.EntireRow.Delete(XlDirection.xlUp);
                                Marshal.ReleaseComObject(range);
                                range = null;

                            }
                        }



                    }
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;

                }

                Marshal.ReleaseComObject(sheet);

            }

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //xlApp.Visible = true;         

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

        }




        public static void RemoveSheet(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            List<String> MasterAssemblyList = componentList;

            if (MasterAssemblyList == null || MasterAssemblyList.Count == 0)
            {
                Utlity.Log("RemoveSheet: " + "Remove Component List is Empty", logFilePath);
                return;
            }

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utlity.Log("xlApp is NULL", logFilePath);
                return;

            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            Utlity.Log("Removing Variable Sheets, If User does Not Need It...", logFilePath);
            List<Microsoft.Office.Interop.Excel._Worksheet> sheetList = new List<_Worksheet>();
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                //Utlity.Log(sheet.Name, logFilePath);
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        //Utlity.Log("Skipping Custom LTC Sheet: " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                if (MasterAssemblyList.Contains(sheet.Name) == false)
                {
                    Utlity.Log("Going to Delete Sheet: " + sheet.Name, logFilePath);
                    sheetList.Add(sheet);
                }
                else
                {
                    Marshal.ReleaseComObject(sheet);
                }
            }

            try
            {
                xlApp.DisplayAlerts = false;
                for (int i = 0; i < sheetList.Count; i++)
                {
                    sheetList[i].Select();
                    sheetList[i].Activate();
                    sheetList[i].Delete();
                    Marshal.ReleaseComObject(sheetList[i]);
                    sheetList[i] = null;
                }
                xlApp.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                Utlity.Log("Deleting Sheets - Exception: " + ex.Message, logFilePath);
                Utlity.Log("Deleting Sheets - Exception: " + ex.StackTrace, logFilePath);
            }

            sheetList.Clear();

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //xlApp.Visible = true;         

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

        }



        // REQUEST - 28 OCT
        private void SyncFeaturesToSolidEdge(String xlFile, String topLineAssembly, String logFilePath)
        {

            String xlFilePath = xlFile;
            Utlity.Log("SolidEdgeSetFeatureState: xlFilePath: " + xlFilePath, logFilePath);
            Utlity.Log("SolidEdgeSetFeatureState: readFeaturesFromTemplateExcel: " + xlFilePath, logFilePath);
            ExcelReadFeatures.readFeaturesFromTemplateExcel(xlFilePath, logFilePath);
            Utlity.Log("SolidEdgeSetFeatureState: getFeatureLinesList: " + xlFilePath, logFilePath);
            List<FeatureLine> updatedFsList = ExcelReadFeatures.getFeatureLinesList();
            try
            {
                Utlity.Log("SolidEdgeSetFeatureState: setFeatures: " + System.DateTime.Now.ToString(), logFilePath);
                Thread myThread = new Thread(() => SolidEdgeSetFeatureState.SolidEdgeFeatureSyncFromExcel(topLineAssembly, updatedFsList, logFilePath));
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

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = e.Result as List<object>;
            if (genericlist == null || genericlist.Count == 0)
            {
                this.progressBar1.Visible = false;
                return;
            }
            String logFilePath = (String)genericlist[1];
            this.DialogResult = DialogResult.OK;
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            listOfFileNamesInSession.Clear();
            MessageBox.Show("SyncTE Completed");

        }

        [STAThread]
        private void ConnectToSolidEdge(String logFilePath)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

            // Connect to running Solid Edge Instance
            application = SE_SESSION.getSolidEdgeSession();
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

        private void ModifyXL(String xlFilePath, String logFilePath)
        {
            Utlity.Log("Updating Template Excel with Latest Values From Solid Edge..", logFilePath, "INFO");
            ExcelDeltaInterface.SaveDeltaToXL(xlFilePath, logFilePath);
            return;
        }

        private void SyncTEDialog_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            About a = new About();
            a.ShowDialog();
        }
    }
}
