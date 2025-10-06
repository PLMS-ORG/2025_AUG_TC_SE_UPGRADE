using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Teamcenter.ClientX;
using System.Collections;
using System.IO;
using DemoAddInTC.services;
using DemoAddInTC.utils;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Client;
using Teamcenter.Soa.Common;
using Teamcenter.Soa.Client.Model.Strong;
using Teamcenter.Services.Strong.Core._2008_06.DataManagement;
//using Teamcenter.Services.Strong.Bom._2008_06.StructureManagement;
using System.Runtime.InteropServices;
//using Teamcenter.Soa.Client.Model;
using System.Diagnostics;
using Creo_TC_Live_Integration.TcDataManagement;
//using Teamcenter.Services.Strong.Administration._2012_09.PreferenceManagement;
//using Teamcenter.Services.Strong.Administration;
using Teamcenter.Services.Strong.Core._2007_09.DataManagement;
using Teamcenter.Services.Strong.Core._2006_03.DataManagement;
using Constants = DemoAddInTC.services.Constants;
//using Tc3.Schemas.Foldermanagement._2018_06;
/// <summary>



namespace DemoAddInTC.se
{
    class TcAdaptor
    {
        public static ModelObject itemModelObject;
        static List<GroupMember> groupMemberList = new List<GroupMember>();

        public static DataManagementService dmService;

        public static SessionService session;

        public static Teamcenter.Soa.Client.Connection tc_Connection;

        public static Teamcenter.ClientX.Session m_session = null;

        public static Teamcenter.Services.Strong.Query.SavedQueryService savedQryServices = null;

        public static SessionService ss = null;

        internal static void get_Session_Log(String logFilePath)
        {
            try
            {
                if (tc_Connection == null)
                {
                    tc_Connection = Teamcenter.ClientX.Session.getConnection();
                }

                ss = SessionService.getService(tc_Connection);
                Teamcenter.Services.Strong.Core._2007_01.Session.GetTCSessionInfoResponse res = ss.GetTCSessionInfo();

                if (res.ExtraInfo.ContainsKey("syslogFile"))
                {
                    String syslogFile_Full_Path = (string)res.ExtraInfo["syslogFile"];

                    Utility.Log("INFO : Syslog file : " + syslogFile_Full_Path, logFilePath);
                }
                else
                {
                    Utility.Log("INFO : Could not find sys log file. ", logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utility.Log("get_Session_Log: " + ex.Message, logFilePath);
                Utility.Log("get_Session_Log: " + ex.StackTrace, logFilePath);
            }
        }

        public static Boolean login(String user, String pwd, String group, String role, String logFilePath)
        {
            try
            {
                bool connect = ConnectToPropertiesFile(logFilePath);
                if (connect == false)
                {
                    Utility.Log("Could not connect to Teamcenter.properties File", logFilePath);
                    return false;
                }
                String serverHost = TCPropertyReader.get(Constants.TC_SERVER_HOST);
                Utility.Log("serverHost: " + serverHost, logFilePath);
                //String serverHost = Constants.TC_SERVER_HOST;
                if (serverHost == null || serverHost.Equals(""))
                {
                    Utility.Log("serverHost is Empty in Teamcenter.properties File", logFilePath);
                    return false;
                }
                else
                {
                    Utility.Log(serverHost, logFilePath);
                }
                if (m_session == null)
                {
                    Utility.Log("Logging into TC 4T", logFilePath);
                    Teamcenter.ClientX.Session session = new Teamcenter.ClientX.Session(serverHost);
                    m_session = session;
                }

                // Establish a session with the Teamcenter Server

                Object isLogInSuccessFlag = "";

                if (group.CompareTo("") == 0 || String.IsNullOrEmpty(group))
                {

                    isLogInSuccessFlag = m_session.login(user, pwd, "", "", "", logFilePath);

                }
                else
                {
                    Utility.Log("serverHost: " + serverHost, logFilePath);
                    Utility.Log("user: " + user, logFilePath);
                    Utility.Log("group: " + group, logFilePath);
                    Utility.Log("role: " + role, logFilePath);
                    isLogInSuccessFlag = m_session.login(user, pwd, group, role, "", logFilePath);
                }
                if (isLogInSuccessFlag == null)
                {
                    return false;
                }
                else
                {
                    TcAdaptor_Init(logFilePath);
                    return true;
                }
            }

            catch (Exception e)
            {
                //MessageBox.Show(e.ToString());
                MessageBox.Show(e.Message);
                MessageBox.Show(e.StackTrace);
                Utility.Log("Message: " + e.Message, logFilePath);
                Utility.Log("InnerException: " + e.InnerException.ToString(), logFilePath);
                Utility.Log("StackTrace: " + e.StackTrace, logFilePath);
                return false;
            }
        }

        public static bool ConnectToPropertiesFile(String logFilePath)
        {
            
            String path = Path.Combine(utils.Utlity.getExecutingPath(),Constants.TEAMCENTER_PROPERTIES_FILE);
            Utility.Log("ConnectToPropertiesFile: " + path, logFilePath);

            if (File.Exists(path) == true)
            {
                TCPropertyReader.reload(path);
            }
            else
            {
                return false;
            }
            return true;
        }

        public static void TcAdaptor_Init(String logFilePath)
        {
            try
            {
                if (dmService == null)
                {
                    dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());
                }
                if (session == null)
                {

                    session = SessionService.getService(Teamcenter.ClientX.Session.getConnection());
                    setObjectPolicy_General(logFilePath);
                }

                if (tc_Connection == null)
                {
                    tc_Connection = Teamcenter.ClientX.Session.getConnection();
                }

                if (savedQryServices == null)
                { 
                    savedQryServices = Teamcenter.Services.Strong.Query.SavedQueryService.getService(tc_Connection);
                }
            }
            catch (Exception ex)
            {
                Utility.Log("TcAdaptor_Init, Exception:" + ex.Message, logFilePath);
                Utility.Log("TcAdaptor_Init, Exception:" + ex.StackTrace, logFilePath);

            }
        }

        public static void setObjectPolicy_General(String logFilePath)
        {
            String[] props = new String[11];


            props[0] = "uid";

            //props[1] = "cd4InstanceIdCp";

            props[1] = "release_statuses";

            props[2] = "release_status_list";

            props[3] = "current_revision_id";

            props[4] = "checked_out";

            props[5] = "isObsolete";

            props[6] = "revision_list";

            props[7] = "item_id";

            props[8] = "item_revision_id";

            props[9] = "IMAN_specification";

            props[10] = "based_on";
            
            ObjectPropertyPolicy policy = new ObjectPropertyPolicy();

            String Revision = Utlity.getPropertyValue("ITEM_REVISION_TYPE", logFilePath); //"P7MCPgPartRevision"
            Utility.Log("setObjectPolicy: ITEM_REVISION_TYPE: " + Revision, logFilePath);
            policy.AddType(new PolicyType(Revision, props));
            
            String Item = Utlity.getPropertyValue("ITEM_TYPE", logFilePath); //"P7MCPgPart"
            Utility.Log("setObjectPolicy: ITEM_TYPE: " + Item, logFilePath);
            policy.AddType(new PolicyType(Item, new string[] { "object_type", "object_name", "item_id", "bom_view_tags", "item_revision_id", "IMAN_reference", "IMAN_specification", "revision_list", "bom_view_tags" }));

            policy.AddType(new PolicyType("BOMLine", new String[] { "bl_bomview_rev", "bl_bomview", "CD4InstanceId", "cd4InstanceId", "bl_line_name", "bl_sequence_no", "bl_plmxml_abs_xform", "bl_all_notes", "bl_quantity", "SE ObjectID", "ps_children", "bl_rev_ps_children", "bl_all_child_lines", "bl_child_lines" }));

            policy.AddType(new PolicyType("BOMWindow", new String[] { "is_packed_by_default" }));

            policy.AddType("RevisionRule", new String[] { "object_name", "Rule_date" });

            policy.AddType("RevisionRuleInfo", new String[] { "object_name", "Rule_date" });

            policy.AddType("ReleaseStatus", new String[] { "Name", "name", "object_name", "name" });

            policy.AddType("release_status_list", new String[] { "object_name", "name" });

            //policy.AddType(new PolicyType("Dataset", new String[] { "object_string", "cd4UploadTime", "object_desc", "ref_list", "checked_out", "dataset_type", "datasettype_name", "object_type", "original_file_name" }));

            policy.AddType(new PolicyType("Dataset", new String[] { "object_string", "object_desc", "ref_list", "checked_out", "dataset_type", "datasettype_name", "object_type", "original_file_name" }));

            policy.AddType(new PolicyType("DatasetType", new String[] { "cd4UploadTime", "Dataset_type", "datasettype_name", "type_name", "original_file_name", "ref_list" }));

            policy.AddType(new PolicyType("ImanFile", new String[] { "original_file_name" }));

            //added by pragati for group,role,group members : April 2024
            //policy.AddType("GroupMember", new String[] { "group", "the_group", "list_of_role","role" });

            //policy.AddType("Group", new String[] { "display_name", "Display_name", "list_of_role", "List_of_role" });

            //policy.AddType("Role", new String[] { "display_name", "role_name" });


            session.SetObjectPropertyPolicy(policy);
        }


        public static Dataset createDataSet2(ItemRevision item_Revision_ModelObject, string dataSetName)
        {
            Teamcenter.Services.Strong.Core._2008_06.DataManagement.DatasetProperties2[] props = new Teamcenter.Services.Strong.Core._2008_06.DataManagement.DatasetProperties2[1];

            props[0] = new Teamcenter.Services.Strong.Core._2008_06.DataManagement.DatasetProperties2();

            props[0].ClientId = dataSetName;

            //props[0].Type = Constants.tc_DataSet_Type;
            props[0].Type = "MSExcelX";

            props[0].Name = dataSetName;

            props[0].Description = dataSetName;

            props[0].Container = item_Revision_ModelObject;

            //props[0].RelationType = Constants.tc_Dataset_Attach_Type;
            props[0].RelationType = "IMAN_specification";

            Teamcenter.Services.Strong.Core._2006_03.DataManagement.CreateDatasetsResponse response_CreateDataset = dmService.CreateDatasets2(props);

            Dataset dataSetModelObj = response_CreateDataset.Output[0].Dataset;
            
            dmService.RefreshObjects(new ModelObject[] { dataSetModelObj });

            dmService.LoadObjects(new string[] { dataSetModelObj.Uid });

            return dataSetModelObj;
        }

        public static void attachFileToDataSet(ModelObject dsMo, string fileWithPath, string dataSetType, string dataSetName, string dataSetDescription, string nameRefType, string logFilePath)
        {
            try
            {
                if (dsMo == null)
                {
                    Utility.Log("Dataset Model Object is Empty..", logFilePath);
                    return;
                }
                FileInfo file1 = new FileInfo(fileWithPath);

                // Create a file to associate with dataset
                Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo[] fileInfo = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo[1];
                fileInfo[0] = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo();
                fileInfo[0].ClientId = "createDataSetAndAttachFile";
                fileInfo[0].FileName = fileWithPath;// file1.Name;
                fileInfo[0].NamedReferencedName = nameRefType;
                fileInfo[0].IsText = false;
                fileInfo[0].AllowReplace = false;
                //Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo[] fileInfos = { fileInfo };

                Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData[] inputData = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData[1];
                inputData[0] = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData();
                inputData[0].Dataset = dsMo;
                inputData[0].CreateNewVersion = false;
                inputData[0].DatasetFileInfos = fileInfo;

                if (inputData == null)
                {
                    Utility.Log("Input data is empty in attach file to dataset function", logFilePath);
                    return;
                }


                Stream fs1 = new FileStream(fileWithPath, FileMode.Open);

                Stream[] fsArray = new Stream[1];
                fsArray[0] = fs1;
                long[] StreamLength = new long[1];
                StreamLength[0] = fs1.Length;

                String fscHost = TCPropertyReader.get(Constants.TC_FSC_HOST);
                if (fscHost == null || fscHost.Equals(""))
                {
                    Utility.Log("Fsc host string is empty.", logFilePath);
                    return;
                }
                else
                {
                }

                String[] url = { fscHost };

                

                Teamcenter.Soa.Client.FileManagementUtility fileManag_Utility = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection());
                ServiceData serviceResponse = fileManag_Utility.PutFiles(inputData);

                if (serviceResponse.sizeOfPartialErrors() > 0)
                {
                    StringBuilder sb = new StringBuilder();

                    for (int i = 0; i < serviceResponse.sizeOfPartialErrors(); i++)
                    {

                        for (int ii = 0; ii < serviceResponse.GetPartialError(i).Messages.Length; ii++)
                        {
                            String error = serviceResponse.GetPartialError(i).Messages[ii].ToString();

                            sb.Append(error + ",");
                        }
                    }

                    Utility.Log(sb.ToString(),logFilePath);
                }

                fileManag_Utility.Term();

                Utility.Log("Size of created object : " + serviceResponse.sizeOfCreatedObjects(), logFilePath);
                /*** refreshing created object***/

                ModelObject[] mo = new ModelObject[serviceResponse.sizeOfCreatedObjects()];

                for (int i = 0; i < serviceResponse.sizeOfCreatedObjects(); i++)
                {
                    mo[i] = serviceResponse.GetCreatedObject(i);
                }
                if (dmService != null)
                {
                    dmService.RefreshObjects(mo);
                }

            }

            catch (Exception e)
            {

                Utility.Log("AttachFileToDataset: " + e.Message, logFilePath);
                Utility.Log("AttachFileToDataset: " + e.StackTrace, logFilePath);
                Utility.Log("AttachFileToDataset: " + e.InnerException, logFilePath);
            }
        }
        //public static void attachFileToDataSet(ModelObject dsMo, string fileWithPath, string dataSetType, string dataSetName, string dataSetDescription, string nameRefType
        //    , String logFilePath)
        //{
        //    try
        //    {
        //        if (dsMo == null)
        //        {
        //            // return false;
        //        }
        //        FileInfo file1 = new FileInfo(fileWithPath);

        //        // Create a file to associate with dataset
        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo fileInfo = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo();
        //        fileInfo.ClientId = "createDataSetAndAttachFile";
        //        fileInfo.FileName = file1.Name;
        //        fileInfo.NamedReferencedName = nameRefType;
        //        fileInfo.IsText = false;
        //        fileInfo.AllowReplace = false;
        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo[] fileInfos = { fileInfo };

        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData inputData = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData();
        //        inputData.Dataset = dsMo;
        //        inputData.CreateNewVersion = false;
        //        inputData.DatasetFileInfos = fileInfos;


        //        Stream fs1 = new FileStream(fileWithPath, FileMode.Open);

        //        Stream[] fsArray = new Stream[1];
        //        fsArray[0] = fs1;
        //        long[] StreamLength = new long[1];
        //        StreamLength[0] = fs1.Length;

        //        String fscHost = TCPropertyReader.get(Constants.TC_FSC_HOST);
        //        if (fscHost == null || fscHost.Equals(""))
        //        {
        //            Utility.Log("TC_FSC_HOST: " + " is EMPTY", logFilePath);
        //            return;
        //        }


        //        Utility.Log("TC_FSC_HOST: " + TCPropertyReader.get(Constants.TC_FSC_HOST), logFilePath);
        //        String[] url = { TCPropertyReader.get(Constants.TC_FSC_HOST) };

        //        FSCStreamingUtility fsc = new FSCStreamingUtility(Teamcenter.ClientX.Session.getConnection(), "", url, url);



        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData[] inputs = { inputData };

        //        ServiceData sd = fsc.Upload(inputs, fsArray, StreamLength);

        //        if (sd.sizeOfPartialErrors() > 0)
        //        {
        //            Utility.Log("Partial Errors : " + sd.sizeOfPartialErrors().ToString(), logFilePath);


        //            string error = "";

        //            for (int j = 0; j < sd.sizeOfPartialErrors(); j++)
        //            {

        //                for (int k = 0; k < sd.GetPartialError(j).Messages.Count(); k++)
        //                {
        //                    error = error + "." + sd.GetPartialError(j).Messages[k];
        //                }
        //            }

        //            Utility.Log("error : " + error, logFilePath);
        //        }

        //        if (sd.sizeOfCreatedObjects() > 0 || sd.sizeOfUpdatedObjects() > 0 ||
        //            sd.sizeOfPlainObjects() >0 || sd.sizeOfDeletedObjects() > 0)
        //        {
        //            Utility.Log("Size of Created Objects : " + sd.sizeOfCreatedObjects().ToString(), logFilePath);
        //            Utility.Log("Size of Updated Objects : " + sd.sizeOfUpdatedObjects().ToString(), logFilePath);
        //            Utility.Log("Size of Plain Objects : " + sd.sizeOfPlainObjects().ToString(), logFilePath);
        //            Utility.Log("Size of Deleted Objects : " + sd.sizeOfDeletedObjects().ToString(), logFilePath);
        //            //  OutputBlock.Inlines.Add("DEBUG - Error in Upload" + "\n");
        //            //  fsc.Term();
        //        }


        //        fsc.Term();
        //        fs1.Close();
        //    }

        //    catch (Exception e)
        //    {
        //        MessageBox.Show("Error attaching the file inside the dataset" + e);
        //    }
        //}

        public static ModelObject isDataSetAvailable(ModelObject itemRevisionModelObject, String typeRef, String datasetType)
        {
            ItemRevision itemRev = (ItemRevision)itemRevisionModelObject;
            if (itemRev == null)
            {

                return null;
            }

            ModelObject[] datasetMoArray = itemRev.IMAN_specification;


            if (datasetMoArray != null && datasetMoArray.Length > 0)
            {
                foreach (ModelObject ds in datasetMoArray)
                {
                    ModelObject dataSetMo = (Dataset)ds;

                    String dsUID = dataSetMo.Uid;
                    
                    String[] arr = new String[] { dsUID };

                    ServiceData sData = dmService.LoadObjects(arr);

                    Dataset dsMObj = (Dataset)sData.GetPlainObject(0);

                    String dsTypeNameInTc = dsMObj.GetPropertyDisplayableValue("object_type"); //Verify with below line


                    if (dataSetMo == null)
                    {

                        return null;
                    }
                    else
                    {
                        if (dsTypeNameInTc.CompareTo(datasetType) == 0)
                        {
                            dmService.RefreshObjects(new ModelObject[] { dsMObj });
                            return dsMObj;
                        }
                    }
                }
            }
            return null;
        }

        public static ModelObject removeDataSetReference(ModelObject dataSetMo, ModelObject itemRevisionModelObject, String typeRef, String dataSetType)
        {
            //  ModelObject dataSetMo = isDataSetAvailable(itemRevisionModelObject, typeRef, dataSetType);

            /*** File details to remove ***/

            Teamcenter.Services.Strong.Core._2007_09.DataManagement.NamedReferenceInfo[] refFileInfo = new Teamcenter.Services.Strong.Core._2007_09.DataManagement.NamedReferenceInfo[1];

            refFileInfo[0] = new Teamcenter.Services.Strong.Core._2007_09.DataManagement.NamedReferenceInfo();

            refFileInfo[0].DeleteTarget = false;

            // refFileInfo[0].Type = "P7MCpkg";

            //refFileInfo[0].Type = "CD4CreoPackage";

            refFileInfo[0].Type = typeRef;

            //   refFileInfo[0].TargetObject = moDsRefList[0]; // To change


            /*** Named reference details ***/
            Teamcenter.Services.Strong.Core._2007_09.DataManagement.RemoveNamedReferenceFromDatasetInfo[] inputInfo = new Teamcenter.Services.Strong.Core._2007_09.DataManagement.RemoveNamedReferenceFromDatasetInfo[1];


            inputInfo[0] = new Teamcenter.Services.Strong.Core._2007_09.DataManagement.RemoveNamedReferenceFromDatasetInfo();

            inputInfo[0].ClientId = "DataRemove1";

            inputInfo[0].Dataset = (Dataset)dataSetMo;//.GetProperty("ref_list").ModelObjectValue;

            inputInfo[0].NrInfo = refFileInfo;


            /*** removing named reference from dataset***/

            // DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            ServiceData outputResponse = dmService.RemoveNamedReferenceFromDataset(inputInfo);

            //   dmService.GetRevNRAttachDetails(

            if (outputResponse.sizeOfPartialErrors() > 0)
            {
                // CreoToTc.Utils.Logger.Log(Constants.MAIN, "removeDataSetReference: " + outputResponse.ToString());

                if (dataSetMo != null)
                {
                    return (Dataset)dataSetMo;
                }
                else
                {
                    return null;
                }
            }

            return (Dataset)dataSetMo;
        }


        public static void logout(String logFilePath)
        {
            if (session == null)
            {
                Utility.Log("TC logout:" + " session is null/empty",logFilePath);
                return;
            }
            try
            {                
                session.Logout();
                //session = null;
                //m_session = null;
            }
            catch (Exception ex)
            {
                Utility.Log("TC logout:" + " Unable to logout from teamcenter..", logFilePath);

            }
        }

        public static void uploadExcelToTC(String user, String pwd, String group, String role, string outputXLfileName, string logFilePath)
        {
            //TcAdaptor Tc = new TcAdaptor();
           
            //bool logIn_Success = TcAdaptor.login(user, pwd, group,role, logFilePath);
            //TcAdaptor.TcAdaptor_Init(logFilePath);
            try
            {
                SEECAdaptor.LoginToTeamcenter(logFilePath);
            }
            catch (Exception)
            {

            }
            

            String itemID = System.IO.Path.GetFileNameWithoutExtension(outputXLfileName);
            String topLevelAssemblyName = System.IO.Path.ChangeExtension(outputXLfileName, ".asm");
            String RevID = SEECAdaptor.getRevisionID(topLevelAssemblyName);

            //if (logIn_Success == false)
            //{
                //MessageBox.Show("Teamcenter Login failed");
                //return;
            //}
            //else
            //{
                //MessageBox.Show("Teamcenter Login Success");
            //}

            // DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            //DownloadDatasetNamedReference.setObjectPolicy();
            Utility.Log("uploadExcelToTC: getItemRevisionQuery..", logFilePath);
            ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID, logFilePath);

            if (itemRevMO == null)
            {
                Utility.Log("uploadExcelToTC: item REV model Object is NULL/Empty", logFilePath);
                return;

            }

            Utility.Log("uploadExcelToTC: isDataSetAvailable..", logFilePath);
            ModelObject dataSetMo = isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");

            if (dataSetMo == null)
            {
                Utility.Log("uploadExcelToTC: createDataSet2..", logFilePath);
                Dataset dsCreated = createDataSet2(itemRevMO, itemID);
                Utility.Log("uploadExcelToTC: attachFileToDataSet..", logFilePath);
                attachFileToDataSet(dsCreated, outputXLfileName, "MS ExcelX", itemID, "Test", "excel", logFilePath);
            }
            else
            {
                Utility.Log("uploadExcelToTC: removeDataSetReference..", logFilePath);
                removeDataSetReference(dataSetMo, itemRevMO, "excel", "MS ExcelX");
                Utility.Log("uploadExcelToTC: attachFileToDataSet..", logFilePath);
                if (File.Exists(outputXLfileName) == true)
                {
                    attachFileToDataSet(dataSetMo, outputXLfileName, "MS ExcelX", itemID, "Test", "excel", logFilePath);
                } else
                {
                    Utility.Log("outputXLfileName is not found: " + outputXLfileName, logFilePath);
                }

                Utility.Log("uploadExcelToTC: checkInModelObject..", logFilePath);
                checkInModelObject(dataSetMo);
            }
              
            //logout(logFilePath);
            //Utility.Log("uploadExcelToTC: Logging out of Teamcenter..", logFilePath);

        }


        //=========================================================================================
        internal static ModelObject checkOutModelObject(ModelObject mo, string logFilePath)
        {
            try
            {
                ReservationService rs = ReservationService.getService(Teamcenter.ClientX.Session.getConnection());

                if (mo != null)
                {
                    ModelObject[] mObjectArray = new ModelObject[] { mo };

                    ServiceData sd = rs.Checkout(mObjectArray, "check out", "");
               
                    Utility.Log("checkOutModelObject: size of updated object :" + sd.sizeOfUpdatedObjects(), logFilePath);

                    if (sd.sizeOfUpdatedObjects() > 0)
                    {
                        ModelObject returnMO = sd.GetUpdatedObject(0);
                        Utility.Log("Successfully checked out",logFilePath);
                        return returnMO;
                    }
                    else
                    {
                        //   MessageBox.Show("Dataset is already checked out");
                        int numPartialError = sd.sizeOfPartialErrors();
                        if(numPartialError > 0 )
                        Utility.Log("checkOutModelObject : Servicedata Partial error:", logFilePath);
                        for (int iError = 0; iError < numPartialError; ++iError)
                        {
                            Utility.Log(iError + ". " + sd.GetPartialError(iError).ToString(),logFilePath);
                        }

                        Utility.Log("Failes: checked out ",logFilePath);
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            } 
            catch (Exception ex)
            {
                Utility.Log("Exception in checkOutModelObject" + ex.ToString(),logFilePath);
                Utility.Log("Stack trace :" + ex.StackTrace, logFilePath);
                Utility.Log("Inner Exception :" + ex.InnerException, logFilePath);
                return null;
            }
        }

        internal static bool checkInModelObject(ModelObject mo)
        {
         
            bool isCheckedIn = false;
            try
            {

                ReservationService rs = ReservationService.getService(Teamcenter.ClientX.Session.getConnection());

                if (mo != null)
                {
                    ModelObject[] mObjectArray = new ModelObject[] { mo };

                    ServiceData sd = rs.Checkin(mObjectArray);

                    if (sd.sizeOfUpdatedObjects() > 0)
                    {

                        isCheckedIn = true;
                    }
                    else
                    {

                        isCheckedIn = false;
                    }

                }
                else
                {

                    isCheckedIn = false;
                }
            }
            catch (Exception e)
            {

            }
            return isCheckedIn;

        }

        public static void setIRProperty(ModelObject ModelObj, String value, String name, String logFilePath)
        {
            //ModelObject[] dataSets = new ModelObject[] { ModelObj };
            //Hashtable PropValues = new Hashtable();
            //PropValues.Add("object_desc",value);
            try
            {
                Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo[] datasetPropery = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo[1];
                datasetPropery[0] = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo();
                datasetPropery[0].Object = ModelObj;

                datasetPropery[0].VecNameVal = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.NameValueStruct1[1];
                List<String> Prop_RealNameList = new List<String> { name };
                for (int i = 0; i < 1; i++)
                {
                    datasetPropery[0].VecNameVal[i] = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.NameValueStruct1();
                    Utility.Log("setIRProperty Name : " + Prop_RealNameList[i], logFilePath);
                    datasetPropery[0].VecNameVal[i].Name = Prop_RealNameList[i];
                    Utility.Log("setIRProperty value : " + value, logFilePath);
                    datasetPropery[0].VecNameVal[i].Values = new String[] { value };
                }

                if (dmService == null)
                {
                    Utility.Log("setIRProperty dmService is NULL: ", logFilePath);
                    return;
                }

                Teamcenter.Services.Strong.Core._2010_09.DataManagement.SetPropertyResponse resp = dmService.SetProperties(datasetPropery, new String[0]);

                if (resp.Data.sizeOfPartialErrors() > 0)
                {
                    Utility.Log("setIRProperty partial Error count: " + resp.Data.sizeOfPartialErrors(), logFilePath);
                }

                if (resp.Data.sizeOfUpdatedObjects() > 0)
                {
                    Utility.Log("setIRProperty updated object count: " + resp.Data.sizeOfUpdatedObjects(), logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utility.Log("setIRProperty Exception: " + ex.Message, logFilePath);
                Utility.Log("setIRProperty Exception: " + ex.StackTrace, logFilePath);
                return;

            }

        }

        public static void setProperty(ModelObject ModelObj, String value, String logFilePath)
        {
            ModelObject[] dataSets = new ModelObject[] {ModelObj};
            //Hashtable PropValues = new Hashtable();
            //PropValues.Add("object_desc",value);

            Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo[] datasetPropery = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo[1];
            datasetPropery[0] = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo();
            datasetPropery[0].Object = ModelObj;

            datasetPropery[0].VecNameVal = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.NameValueStruct1[2];
            List<String> Prop_RealNameList = new List<String> { "object_desc", "object_name"};
            for (int i = 0 ; i < 2; i ++) {
                datasetPropery[0].VecNameVal[i] = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.NameValueStruct1();
                datasetPropery[0].VecNameVal[i].Name = Prop_RealNameList[i];
                datasetPropery[0].VecNameVal[i].Values = new String[] { value };
            }

            Teamcenter.Services.Strong.Core._2010_09.DataManagement.SetPropertyResponse resp = dmService.SetProperties(datasetPropery, new String[0]);

            if (resp.Data.sizeOfPartialErrors() > 0)
            {
                Utility.Log("sizeOfPartialErrors: ", logFilePath);
                Utility.Log("sizeOfPartialErrors: " + resp.Data.sizeOfPartialErrors(), logFilePath);
               
                for (int j = 0; j < resp.Data.sizeOfPartialErrors(); j++)
                {
                    string error = "";

                    for (int k = 0; k < resp.Data.GetPartialError(j).Messages.Count(); k++)
                    {
                        error = error + "." + resp.Data.GetPartialError(j).Messages[k];
                    }

                    Utility.Log("error: " + error, logFilePath);
                }
            }

            if (resp.Data.sizeOfUpdatedObjects() > 0)
            {
                Utility.Log("sizeOfUpdatedObjects: ", logFilePath);
                Utility.Log("setProperty: " + resp.Data.sizeOfUpdatedObjects(), logFilePath);

                for (int j = 0; j < resp.Data.sizeOfUpdatedObjects(); j++)
                {
                    string Uid = "";

                    for (int k = 0; k < resp.Data.sizeOfUpdatedObjects(); k++)
                    {
                        Uid = Uid + ":::" + resp.Data.GetUpdatedObject(k).Uid;
                        int len = resp.Data.GetUpdatedObject(k).PropertyNames.Length;

                        for (int kk =0; kk < len; kk++ )
                        {
                            String name = resp.Data.GetUpdatedObject(k).PropertyNames[kk];
                            string value11 = resp.Data.GetUpdatedObject(k).GetPropertyDisplayableValue(name);
                            Utility.Log("Updated Object Name : " + name, logFilePath);
                            Utility.Log("Updated Object Value : " + value11, logFilePath);
                        }
                    }

                    Utility.Log("Updated Object UID/s: " + Uid, logFilePath);
                }
            }

        }


        public static void PostCloneCleanUpExcelDataSet(string itemID, string RevID, string logFilePath)
        {
           ModelObject itemRevObj = DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID, logFilePath);


           if (itemRevObj == null)
                {

                    Utility.Log("Item Rev Model Object is NULL: " + itemID, logFilePath);
                    return ;
                }

                
                {
                    ModelObject[] dataSets = DownloadDatasetNamedReference.getAllDataSet(itemRevObj, logFilePath);
                    if (dataSets != null)
                    {

                        foreach (ModelObject dsMo in dataSets)
                        {
                            Dataset ds = (Dataset)dsMo;

                            String dsUID = ds.Uid;
                            //String object_type = dsMo.GetPropertyDisplayableValue("object_type");


                            DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

                            String[] arr = new String[] { dsUID };

                            ServiceData sData = dmService.LoadObjects(arr);
                            

                            Dataset dsMObj = (Dataset)sData.GetPlainObject(0);
                            String object_Type = dsMObj.Object_type;                            
                            Utility.Log("DsType:" + object_Type, logFilePath);
                            
                            if (object_Type.Equals("MSExcelX", StringComparison.OrdinalIgnoreCase) == true)
                            {
                            Utility.Log("itemID to set on Dataset :" + itemID, logFilePath);
                            setProperty(dsMObj, itemID, logFilePath);
                            }
                        }

                    }
                    else
                    {
                        Utility.Log("DataSet is empty for ItemID/RevID :" + itemID + "/" + RevID, logFilePath);
                    }
                }
        }
    }
}
