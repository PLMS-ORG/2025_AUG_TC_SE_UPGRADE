using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _2DAutomaticUpdate
{
    class SyncDwg
    {
        public static List<string> listOfFileNamesInSession = new List<string>();
        public static void SyncDwgs(String itemID, String revID, String logFilePath)
        {
            Utility.Log("Connecting to Solid Edge..", logFilePath);
            
            //Finding drawings of current assembly
            SolidEdgeFramework.Application Seapplication = null;
            SolidEdgeFramework.SolidEdgeDocument Sedocument = null;
            Seapplication = LTC_SEEC.getSolidEdgeSession();
            //if (Seapplication == null)
            //{
            //    Utility.Log("Solid Edge Application is NULL",logFilePath);
            //    return;
            //}
            //Sedocument = (SolidEdgeFramework.SolidEdgeDocument)Seapplication.ActiveDocument;

            //if (Sedocument == null)
            //{
            //    Utility.Log("Solid Edge Document is NULL",logFilePath);
            //    return;
            //}
            try
            {
                SolidEdgeFramework.SolidEdgeTCE objSEEC = Seapplication.SolidEdgeTCE;
                //SolidEdgeFramework.PropertySets propertySets = (SolidEdgeFramework.PropertySets)Sedocument.Properties;
                //SolidEdgeFramework.Properties projectInformation = (SolidEdgeFramework.Properties)propertySets.Item(5);
                //SolidEdgeFramework.Property revision = (SolidEdgeFramework.Property)projectInformation.Item(2);
                //string revisionValue = revision.get_Value().ToString();
                //SolidEdgeFramework.Property documentNumber = (SolidEdgeFramework.Property)projectInformation.Item(1);
                //string documentNumberValue = documentNumber.get_Value().ToString();
                int NoOfComponents = 0;
                System.Object ListOfItemRevIds = null, ListOfFileSpecs = null;
                objSEEC.GetBomStructure(itemID, revID, SolidEdgeConstants.RevisionRuleType.LatestRevision.ToString(), true, out NoOfComponents,
                        out ListOfItemRevIds, out ListOfFileSpecs);
                Dictionary<string, string> itemAndRevIds = new Dictionary<string, string>();
                itemAndRevIds.Add(itemID, revID);
                if (NoOfComponents > 0)
                {
                    System.Object[,] abcd = (System.Object[,])ListOfItemRevIds;
                    for (int i = 0; i <= abcd.GetUpperBound(0); i++)
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
                    Utility.Log("Trying to download " + pair.Key + "/" + pair.Value, logFilePath);
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
                            Utility.Log("Downloaded drawing of " + pair.Key + "/" + pair.Value + " to cache", logFilePath);
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

            // Update the Drafts
           
            try
            {
                SolidEdgeFramework.SolidEdgeTCE objSEEC = (SolidEdgeFramework.SolidEdgeTCE)Seapplication.SolidEdgeTCE;
                Utility.Log("Updating Views in Draft Files....", logFilePath);
                string cachePath = null;
                Utility.Log("SEEC Getting cache path..", logFilePath);
                objSEEC.GetPDMCachePath(out cachePath);
                SolidEdgeUpdateView.SearchDraftFile(cachePath, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("SearchDraftFile: " + ex.Message, logFilePath);
            }

            // Upload the files back to Teamcenter - 17 August 2019
            try
            {
                Utility.Log("Uploading files back to Teamcenter....", logFilePath);

                SolidEdgeFramework.SolidEdgeTCE objSEEC = (SolidEdgeFramework.SolidEdgeTCE)Seapplication.SolidEdgeTCE;
                Seapplication.DisplayAlerts = false;
                objSEEC.SetTeamCenterMode(true);
                String fmsHome = System.Configuration.ConfigurationManager.AppSettings["FMS Home"];
                String userName = System.Configuration.ConfigurationManager.AppSettings["User Name"];
                String password = System.Configuration.ConfigurationManager.AppSettings["Password"];
                String group = System.Configuration.ConfigurationManager.AppSettings["Group"];
                String role = System.Configuration.ConfigurationManager.AppSettings["Role"];
                String URL = System.Configuration.ConfigurationManager.AppSettings["URL"];

                objSEEC.ValidateLogin(userName, password, group, role, URL);
                Utility.Log("SEEC Log in Successful..", logFilePath);


                LTC_SEEC.CheckInSEDocumentsToTeamcenter(logFilePath, listOfFileNamesInSession);
                
            }
            catch (Exception ex)
            {
                Utility.Log("Uploading files back to Teamcenter.... " + ex.Message, logFilePath);                
                return;
            }
         

        }
    }
}
