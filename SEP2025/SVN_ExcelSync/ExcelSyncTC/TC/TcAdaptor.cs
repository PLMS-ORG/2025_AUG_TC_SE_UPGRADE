using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Teamcenter.ClientX;
using System.Collections;
using System.IO;
using ExcelSyncTC.services;
using ExcelSyncTC.utils;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Client;
using Teamcenter.Soa.Common;
using Teamcenter.Soa.Client.Model.Strong;
using Teamcenter.Services.Strong.Core._2008_06.DataManagement;
using Teamcenter.Services.Strong.Bom._2008_06.StructureManagement;
using System.Runtime.InteropServices;
using Teamcenter.Soa.Client.Model;
using System.Diagnostics;
using Teamcenter.Services.Strong.Core._2007_09.DataManagement;
using Teamcenter.Services.Strong.Core._2006_03.DataManagement;


namespace ExcelSyncTC.TC
{
    class TcAdaptor
    {
        public static ModelObject itemModelObject;
        static List<GroupMember> groupMemberList = new List<GroupMember>();

        public static DataManagementService dmService;

        public static SessionService session;

        public static Teamcenter.Soa.Client.Connection tc_Connection;

        public static Teamcenter.ClientX.Session m_session = null;
        public static Boolean login(String user, String pwd, String group, String role, String logFilePath)
        {
            try
            {
                //bool connect = ConnectToPropertiesFile();
                //if (connect == false)
                //{

                //    return false;
                //}
                //  String serverHost = TCPropertyReader.get(Constants.TC_SERVER_HOST);

                String serverHost = Constants.TC_SERVER_HOST;
                if (serverHost == null || serverHost.Equals(""))
                {
                    return false;
                }
                else
                {
                    Utlity.Log(serverHost, logFilePath);
                }

                Teamcenter.ClientX.Session session = new Teamcenter.ClientX.Session(serverHost);
                m_session = session;

                // Establish a session with the Teamcenter Server

                Object isLogInSuccessFlag = "";

                if (group.CompareTo("") == 0 || String.IsNullOrEmpty(group))
                {

                    isLogInSuccessFlag = session.login(user, pwd, "", "", "", logFilePath);

                }
                else
                {
                    isLogInSuccessFlag = session.login(user, pwd, group, role, "", logFilePath);
                }
                if (isLogInSuccessFlag == null)
                {
                    return false;
                }
                return true;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
        }

        //public static bool ConnectToPropertiesFile()
        //{
        //    String path = Path.Combine(Constants.TEAMCENTER_PROPERTIES_FILE);

        //    if (File.Exists(path) == true)
        //    {
        //        TCPropertyReader.reload(path);
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //    return true;
        //}

        public static void TcAdaptor_Init()
        {

            dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            session = SessionService.getService(Teamcenter.ClientX.Session.getConnection());

            tc_Connection = Teamcenter.ClientX.Session.getConnection();

            setObjectPolicy_General();
        }

        private static void setObjectPolicy_General()
        {
            String[] props = new String[11];


            props[0] = "uid";

            props[1] = "cd4InstanceIdCp";

            props[2] = "release_statuses";

            props[3] = "release_status_list";

            props[4] = "current_revision_id";

            props[5] = "checked_out";

            props[6] = "isObsolete";

            props[7] = "revision_list";

            props[8] = "item_id";

            props[9] = "item_revision_id";

            props[10] = "IMAN_specification";

            ObjectPropertyPolicy policy = new ObjectPropertyPolicy();

            policy.AddType(new PolicyType("ItemRevision", props));

            policy.AddType(new PolicyType("Item", new string[] { "object_type", "object_name", "item_id", "bom_view_tags", "item_revision_id", "IMAN_reference", "IMAN_specification", "revision_list", "bom_view_tags" }));

            policy.AddType(new PolicyType("BOMLine", new String[] { "bl_bomview_rev", "bl_bomview", "CD4InstanceId", "cd4InstanceId", "bl_line_name", "bl_sequence_no", "bl_plmxml_abs_xform", "bl_all_notes", "bl_quantity", "SE ObjectID", "ps_children", "bl_rev_ps_children", "bl_all_child_lines", "bl_child_lines" }));

            policy.AddType(new PolicyType("BOMWindow", new String[] { "is_packed_by_default" }));

            policy.AddType("RevisionRule", new String[] { "object_name", "Rule_date" });

            policy.AddType("RevisionRuleInfo", new String[] { "object_name", "Rule_date" });

            policy.AddType("ReleaseStatus", new String[] { "Name", "name", "object_name", "name" });

            policy.AddType("release_status_list", new String[] { "object_name", "name" });

            policy.AddType(new PolicyType("Dataset", new String[] { "object_string", "cd4UploadTime", "object_desc", "ref_list", "checked_out", "dataset_type", "datasettype_name", "object_type" }));

            policy.AddType(new PolicyType("DatasetType", new String[] { "cd4UploadTime", "Dataset_type", "datasettype_name", "type_name" }));

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

        public static void attachFileToDataSet(ModelObject dsMo, string fileWithPath, string dataSetType, string dataSetName, string dataSetDescription, string nameRefType)
        {
            try
            {
                if (dsMo == null)
                {
                    // return false;
                }
                FileInfo file1 = new FileInfo(fileWithPath);

                // Create a file to associate with dataset
                Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo fileInfo = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo();
                fileInfo.ClientId = "createDataSetAndAttachFile";
                fileInfo.FileName = file1.Name;
                fileInfo.NamedReferencedName = nameRefType;
                fileInfo.IsText = false;
                fileInfo.AllowReplace = false;
                Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo[] fileInfos = { fileInfo };

                Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData inputData = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData();
                inputData.Dataset = dsMo;
                inputData.CreateNewVersion = false;
                inputData.DatasetFileInfos = fileInfos;


                Stream fs1 = new FileStream(fileWithPath, FileMode.Create);

                Stream[] fsArray = new Stream[1];
                fsArray[0] = fs1;
                long[] StreamLength = new long[1];
                StreamLength[0] = fs1.Length;

                //String fscHost = TCPropertyReader.get(Constants.TC_FSC_HOST);
                //if (fscHost == null || fscHost.Equals(""))
                //{
                //    Utils.Logger.Log(Constants.LOGIN, "fscHost is Empty ");
                //    return;
                //}
                //else
                //{
                //    Utils.Logger.Log(Constants.LOGIN, "serverHost is :" + fscHost);
                //}

                String[] url = { Constants.fsc_Host };

                Teamcenter.Soa.Client.FSCStreamingUtility fsc = new Teamcenter.Soa.Client.FSCStreamingUtility(Teamcenter.ClientX.Session.getConnection(), "", url, url);



                Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData[] inputs = { inputData };
                ServiceData sd = fsc.Upload(inputs, fsArray, StreamLength);
                if (sd.sizeOfPartialErrors() > 0)
                {
                    //  OutputBlock.Inlines.Add("DEBUG - Error in Upload" + "\n");
                    //  fsc.Term();
                }
                fsc.Term();
                fs1.Close();
            }

            catch (Exception e)
            {
                MessageBox.Show("Error attaching the file inside the dataset" + e);
            }
        }

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


        internal static void logout(String logFilePath)
        {
            if (session == null)
            {
                Utlity.Log("TC logout:" + " session is null/empty", logFilePath);
                return;
            }
            try
            {
                session.Logout();
            }
            catch (Exception ex)
            {
                Utlity.Log("TC logout:" + " Unable to logout from teamcenter..", logFilePath);

            }
        }

        public static void uploadExcelToTC(string outputXLfileName, string logFilePath)
        {
            //TcAdaptor Tc = new TcAdaptor();
            bool logIn_Success = TcAdaptor.login("dcproxy", "dcproxy", "Engineering", "Designer", logFilePath);
            SEECAdaptor.LoginToTeamcenter();

            String itemID = System.IO.Path.GetFileNameWithoutExtension(outputXLfileName);
            String RevID = SEECAdaptor.getRevisionID(outputXLfileName);

            if (logIn_Success == false)
            {
                MessageBox.Show("Teamcenter Login failed");
                return;
            }
            else
            {
                //MessageBox.Show("Teamcenter Login Success");
            }

            // DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            //DownloadDatasetNamedReference.setObjectPolicy();
            ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID, logFilePath);


            ModelObject dataSetMo = isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");

            if (dataSetMo == null)
            {
                Dataset dsCreated = createDataSet2(itemRevMO, itemID);
                attachFileToDataSet(dsCreated, outputXLfileName, "MS ExcelX", itemID, "Test", "excel");
            }
            else
            {
                removeDataSetReference(dataSetMo, itemRevMO, "excel", "MS ExcelX");
                attachFileToDataSet(dataSetMo, outputXLfileName, "MS ExcelX", itemID, "Test", "excel");
            }

            logout(logFilePath);

        }
    }
}
