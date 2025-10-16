using ExcelSyncTC.opInterfaces;
using ExcelSyncTC.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelSyncTC.TC;

namespace ExcelSyncTC
{
    public partial class SyncDwg : Form
    {
        public SyncDwg()
        {
            InitializeComponent();

            Trace.Listeners.Add(new ListBoxTraceListener(listBox1));
        }

        private void SyncDwg_Click(object sender, EventArgs e)
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

        public static List<string> listOfFileNamesInSession = new List<string>();
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application application = null;
            List<object> genericlist = e.Argument as List<object>;
            application = (Microsoft.Office.Interop.Excel.Application)genericlist[0];

            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;


            if (xlApp.ActiveWorkbook == null)
            {
                MessageBox.Show("Sync Dwg: Cannot Access Excel Application");
                this.button1.Enabled = true;
                return;
            }
            String xlFilePath = xlApp.ActiveWorkbook.FullName;
            String LogStageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(xlFilePath) + "_" + "SyncTE" + ".txt");
            String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");

            Utlity.Log("SyncTE 2D Started", logFilePath);
            Utlity.Log("Connecting to Solid Edge..", logFilePath, "INFO");
            SE_SESSION.InitializeSolidEdgeSession(logFilePath);

            //Finding drawings of current assembly
            SolidEdgeFramework.Application Seapplication = null;
            SolidEdgeFramework.SolidEdgeDocument Sedocument = null;
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
            try
            {
                SolidEdgeFramework.SolidEdgeTCE objSEEC = Seapplication.SolidEdgeTCE;
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
                        Utlity.Log("itemAndRevId : " + abcd[i, 0].ToString(), logFilePath);
                        if (itemAndRevIds.Keys.Contains(abcd[i, 0].ToString()) == false)
                            itemAndRevIds.Add(abcd[i, 0].ToString(), abcd[i, 1].ToString());
                    }
                }
                Utlity.Log("NoOfComponents: " + NoOfComponents, logFilePath);
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

                        if (filename.Contains(".par"))
                        {
                            System.Object[,] temp = new object[1, 1];
                            application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded part of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                               SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded sheet metal of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded Assembly of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
                        }

                        if (filename.EndsWith(".pwd", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            System.Object[,] temp = new object[1, 1];
                            application.DisplayAlerts = false;
                            //objSEEC.ValidateLogin(userName, password, group, role, URL);
                            objSEEC.DownladDocumentsFromServerWithOptions(pair.Key, pair.Value, filename,
                                SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), "", false, false,
                                1, temp);
                            Utlity.Log("Downloaded weldment of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
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

            // Update the Drafts
            try
            {
                //Utlity.Log("Saving the Changes Done..", logFilePath, "INFO");
                xlApp.ActiveWorkbook.Save();

                Utlity.Log("Updating Views in Draft Files....", logFilePath, "INFO");
                SolidEdgeUpdateView.SearchDraftFile(topLineAssembly, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("SearchDraftFile: " + ex.Message, logFilePath);
                e.Result = "NOK";
             }

            // Upload the files back to Teamcenter - 17 August 2019
            try
            {
                Utlity.Log("Uploading files back to Teamcenter....", logFilePath, "INFO");
                //TcAdaptor.login("dcproxy", "dcproxy", "Engineering", "Designer", logFilePath);
                //TcAdaptor.TcAdaptor_Init();
                SE_SESSION.InitializeSolidEdgeSession(logFilePath);
                SEECAdaptor.LoginToTeamcenter();
                SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath, SyncDwg.listOfFileNamesInSession);
                //TcAdaptor.uploadExcelToTC(xlFilePath, logFilePath);
                //TcAdaptor.logout(logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("Uploading files back to Teamcenter.... " + ex.Message, logFilePath);
                e.Result = "NOK";
            }

            String startString = "SyncTE 2D Started";
            bool parseFlag = Utlity.parseLog(logFilePath, startString);
            if (parseFlag == false) // Something went wrong
            {
                e.Result = "NOK";
            }
            else
                e.Result = "OK";

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result.Equals("OK") == true)
            {
                this.DialogResult = DialogResult.OK;
                this.Dispose();
                
                MessageBox.Show("Sync Dwg completed");
                return;
            }
            else
            {
                this.DialogResult = DialogResult.Cancel;
                this.Dispose();

                MessageBox.Show("Sync Dwg Failed, Check Logs");
                return;
            }
        }

    }
}
