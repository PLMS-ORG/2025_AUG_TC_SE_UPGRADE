using AddToTc.CDAT_BulkUploader;
using Creo_TC_Live_Integration.TcDataManagement;
using CreoToTc.Utils;
using DemoAddInTC.se;
using DemoAddInTC.services;
using DemoAddInTC.utils;
using Log;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Services.Strong.Query._2007_01.SavedQuery;
using Teamcenter.Soa.Client;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;

namespace AddToTc.CTD
{
    internal class TCUtils
    {

        public static Boolean tc_LogIn(String user, String pwd, [Optional] String group)
        {

            try
            {
                String serverHost = Constants.TC_SERVER_HOST;
                if (serverHost == null || serverHost.Equals(""))
                {

                    return false;
                }

                Teamcenter.ClientX.Session session = new Teamcenter.ClientX.Session(serverHost);


                // Establish a session with the Teamcenter Server

                Object isLogInSuccessFlag = "";

                if (group.CompareTo("") == 0 || String.IsNullOrEmpty(group))
                {

                    isLogInSuccessFlag = session.login(user, pwd, "", "", "");

                }
                else
                {
                    isLogInSuccessFlag = session.login(user, pwd, group, "", "");
                }
                if (isLogInSuccessFlag == null)
                {
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                log.writeException(ex, "Tc Login");
                return false;
            }
        }

        //=================================================================================================================
        public static void setProperty(ModelObject ModelObj, String value)
        {
            ModelObject[] dataSets = new ModelObject[] { ModelObj };
            //Hashtable PropValues = new Hashtable();
            //PropValues.Add("object_desc",value);

            Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo[] datasetPropery
                = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo[1];

            datasetPropery[0] = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.PropInfo();
            datasetPropery[0].Object = ModelObj;

            datasetPropery[0].VecNameVal = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.NameValueStruct1[2];
            List<String> Prop_RealNameList = new List<String> { "object_desc", "object_name" };
            for (int i = 0; i < 2; i++)
            {
                datasetPropery[0].VecNameVal[i] = new Teamcenter.Services.Strong.Core._2010_09.DataManagement.NameValueStruct1();
                datasetPropery[0].VecNameVal[i].Name = Prop_RealNameList[i];
                datasetPropery[0].VecNameVal[i].Values = new String[] { value };
            }

            Teamcenter.Services.Strong.Core._2010_09.DataManagement.SetPropertyResponse resp = Tc_Services.dmService.SetProperties(datasetPropery, new String[0]);

            if (resp.Data.sizeOfPartialErrors() > 0)
            {
                log.write(logType.INFO,"setProperty: " + resp.Data.sizeOfPartialErrors());
            }

            if (resp.Data.sizeOfUpdatedObjects() > 0)
            {
                log.write(logType.INFO, "setProperty: " + resp.Data.sizeOfUpdatedObjects());
            }

        }
        //=================================================================================================================
        public static void PostCloneCleanUpExcelDataSet(string itemID, string RevID)
        {
            ModelObject itemRevObj = DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID);


            if (itemRevObj == null)
            {

                log.write(logType.ERROR, "Item Rev Model Object is NULL: " + itemID);
                return;
            }

            {
                ModelObject[] dataSets = DownloadDatasetNamedReference.getAllDataSet(itemRevObj);
                if (dataSets != null)
                {

                    foreach (ModelObject dsMo in dataSets)
                    {
                        Dataset ds = (Dataset)dsMo;

                        String dsUID = ds.Uid;
                        //String object_type = dsMo.GetPropertyDisplayableValue("object_type");

                       // Tc_Services.dmService
                        //DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

                        String[] arr = new String[] { dsUID };

                        ServiceData sData = Tc_Services.dmService.LoadObjects(arr);


                        Dataset dsMObj = (Dataset)sData.GetPlainObject(0);
                        String object_Type = dsMObj.Object_type;
                        //log.write(logType.ERROR, "DsType:" + object_Type);

                        if (object_Type.Equals("MSExcelX", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            setProperty(dsMObj, itemID);
                        }
                    }

                }
                else
                {
                    log.write(logType.ERROR,"DataSet is empty for ItemID/RevID :" + itemID + "/" + RevID);
                }
            }
        }
        //=================================================================================================================
        // SOA Call
        public static void uploadExcelToTC(String user, String pwd, String group, String role, string outputXLfileName)
        {
           
           

            String itemID = System.IO.Path.GetFileNameWithoutExtension(outputXLfileName);
            String topLevelAssemblyName = System.IO.Path.ChangeExtension(outputXLfileName, ".asm");
            String RevID = SEECAdaptor.getRevisionID(topLevelAssemblyName);

           
            log.write(logType.INFO,"uploadExcelToTC: getItemRevisionQuery..");

            ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID);

            if (itemRevMO == null)
            {
                log.write(logType.ERROR,"uploadExcelToTC: item REV model Object is NULL/Empty");
                return;

            }

            log.write(logType.INFO, "uploadExcelToTC: isDataSetAvailable..");
            ModelObject dataSetMo = isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");

            if (dataSetMo == null)
            {
                log.write(logType.INFO, "uploadExcelToTC: createDataSet2..");
                Dataset dsCreated = createDataSet2(itemRevMO, itemID);
                log.write(logType.INFO, "uploadExcelToTC: attachFileToDataSet..");
                attachFileToDataSet(dsCreated, outputXLfileName, "MS ExcelX", itemID, "Test", "excel");
            }
            else
            {
                log.write(logType.INFO, "uploadExcelToTC: removeDataSetReference..");
                removeDataSetReference(dataSetMo, itemRevMO, "excel", "MS ExcelX");
                log.write(logType.INFO, "uploadExcelToTC: attachFileToDataSet..");
                attachFileToDataSet(dataSetMo, outputXLfileName, "MS ExcelX", itemID, "Test", "excel");
                log.write(logType.INFO, "uploadExcelToTC: checkInModelObject..");
                checkInModelObject(dataSetMo);
            }


        }

        //==========================================================================================================
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

                    ServiceData sData = Tc_Services.dmService.LoadObjects(arr);

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
                            Tc_Services.dmService.RefreshObjects(new ModelObject[] { dsMObj });
                            return dsMObj;
                        }
                    }
                }
            }
            return null;
        }

        //==========================================================================================

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

            Teamcenter.Services.Strong.Core._2006_03.DataManagement.CreateDatasetsResponse response_CreateDataset = Tc_Services.dmService.CreateDatasets2(props);

            Dataset dataSetModelObj = response_CreateDataset.Output[0].Dataset;

            Tc_Services.dmService.RefreshObjects(new ModelObject[] { dataSetModelObj });

            Tc_Services.dmService.LoadObjects(new string[] { dataSetModelObj.Uid });

            return dataSetModelObj;
        }

        //==========================================================================================
        // uploading XL to Teamenter using SOA
        //public static void attachFileToDataSet(ModelObject dsMo, string fileWithPath, string dataSetType, string dataSetName, string dataSetDescription, 
        //                                       string nameRefType )
        //{
        //    try
        //    {
        //        if (dsMo == null)
        //        {
        //            log.write(logType.ERROR, "dsMo is NULL: ");
        //            return;

        //        }

        //        log.write(logType.INFO, "fileWithPath is : "+ fileWithPath);

        //        FileInfo file1 = new FileInfo(fileWithPath);

        //        // Create a file to associate with dataset
        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo fileInfo = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo();
        //        fileInfo.ClientId = "createDataSetAndAttachFile";
        //        fileInfo.FileName = file1.Name;
        //        fileInfo.NamedReferencedName = nameRefType;
        //        fileInfo.IsText = false;
        //        fileInfo.AllowReplace = false;
        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.DatasetFileInfo[] fileInfos = { fileInfo };

        //        if (fileInfo == null)
        //        {
        //            log.write(logType.ERROR, "fileInfo is NULL: ");
        //            return;

        //        }

        //        if (File.Exists(fileWithPath) == false)
        //        {
        //            log.write(logType.ERROR, "fileWithPath does not Exist:" + fileWithPath);
        //            return;
        //        }

        //        Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData inputData = new Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData();
        //        if (inputData == null)
        //        {
        //            log.write(logType.ERROR, "inputData is NULL: ");
        //            return;
        //        }

        //        inputData.Dataset = dsMo;
        //        inputData.CreateNewVersion = false;
        //        inputData.DatasetFileInfos = fileInfos;


        //        Stream fs1 = new FileStream(fileWithPath, FileMode.Open);

        //        Stream[] fsArray = new Stream[1];
        //        fsArray[0] = fs1;
        //        long[] StreamLength = new long[1];
        //        StreamLength[0] = fs1.Length;
        //        fs1.Close();

        //        String fscHost = Constants.TC_FSC_HOST;
        //        if (fscHost == null || fscHost.Equals(""))
        //        {
        //            log.write(logType.ERROR,"TC_FSC_HOST: " + " is EMPTY");
        //            return;
        //        }


        //        log.write(logType.INFO, "TC_FSC_HOST: " + Constants.TC_FSC_HOST);
        //        String[] url = { Constants.TC_FSC_HOST };



        //        try
        //        {


        //            Teamcenter.Soa.Client.FileManagementUtility fileManag_Utility = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection());
        //            Teamcenter.Services.Loose.Core._2006_03.FileManagement.GetDatasetWriteTicketsInputData[] inputs = { inputData };

        //            if (inputs == null)
        //            {
        //                log.write(logType.ERROR, "inputs is NULL: ");
        //                return;
        //            }
        //            log.write(logType.INFO, "PutFiles: ");
        //            ServiceData serviceResponse = fileManag_Utility.PutFiles(inputs);

        //            if (serviceResponse == null)
        //            {
        //                log.write(logType.ERROR, "serviceResponse is NULL: ");
        //                return;
        //            }

        //            ModelObject[] mo = new ModelObject[serviceResponse.sizeOfCreatedObjects()];

        //            if (mo == null)
        //            {
        //                log.write(logType.ERROR, "mo is NULL: ");

        //                return;
        //            }

        //            log.write(logType.INFO , "serviceResponse.sizeOfCreatedObjects: " + serviceResponse.sizeOfCreatedObjects());

        //            for (int i = 0; i < serviceResponse.sizeOfCreatedObjects(); i++)
        //            {
        //                mo[i] = serviceResponse.GetCreatedObject(i);
        //                log.write(logType.INFO, "Sucess Service Response: " + i);
        //            }

        //            Tc_Services.dmService.RefreshObjects(mo);
        //            fileManag_Utility.Term();


        //        }

        //        catch (Exception e)
        //        {

        //            log.write(logType.ERROR, "Exception while performing Reference file's upload into Tc .Error :" + e.Message);

        //            log.write(logType.ERROR, "Exception while performing Reference file's upload into Tc .Error :" + e.StackTrace);

        //            log.write(logType.ERROR, "Exception while performing Reference file's upload into Tc .Error :" + e.InnerException);

        //        }
        //    }

        //    catch (Exception e)
        //    {
        //        log.write(logType.ERROR,"Error attaching the file inside the dataset" + e);
        //    }
        //}

        public static void attachFileToDataSet(ModelObject dsMo, string fileWithPath, string dataSetType, string dataSetName, string dataSetDescription, string nameRefType)
        {
            try
            {
                if (dsMo == null)
                {
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
                    log.write(logType.ERROR, "Input data is empty in attach file to dataset function");
                }


                Stream fs1 = new FileStream(fileWithPath, FileMode.Open);

                Stream[] fsArray = new Stream[1];
                fsArray[0] = fs1;
                long[] StreamLength = new long[1];
                StreamLength[0] = fs1.Length;

                String fscHost = "http://aewsrv-tcsr02:4544/";//Constants.TC_FSC_HOST;//TCPropertyReader.get(Constants.TC_FSC_HOST);
                if (fscHost == null || fscHost.Equals(""))
                {
                    log.write(logType.ERROR, "Fsc host string is empty.");
                    return;
                }
                else
                {
                }

                String[] url = { fscHost };

                //  FSCStreamingUtility fsc = new FSCStreamingUtility(Tc_Services.tc_Connection, "", url, url);

                Teamcenter.Soa.Client.FileManagementUtility fileManag_Utility = new Teamcenter.Soa.Client.FileManagementUtility(Tc_Services.tc_Connection);
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

                    log.write(logType.ERROR, sb.ToString());
                }

                fileManag_Utility.Term();

                log.write(logType.INFO, "Size of created object : " + serviceResponse.sizeOfCreatedObjects());
                /*** refreshing created object***/

                ModelObject[] mo = new ModelObject[serviceResponse.sizeOfCreatedObjects()];

                for (int i = 0; i < serviceResponse.sizeOfCreatedObjects(); i++)
                {
                    mo[i] = serviceResponse.GetCreatedObject(i);
                }

                Tc_Services.dmService.RefreshObjects(mo);

            }

            catch (Exception e)
            {

                log.writeException(e, "Attach file to dataset");
            }
        }

        //==========================================================================================
        public static ModelObject removeDataSetReference(ModelObject dataSetMo, ModelObject itemRevisionModelObject, String typeRef, String dataSetType)
        {
            //  ModelObject dataSetMo = isDataSetAvailable(itemRevisionModelObject, typeRef, dataSetType);

            /*** File details to remove ***/

            Teamcenter.Services.Strong.Core._2007_09.DataManagement.NamedReferenceInfo[] refFileInfo =
                new Teamcenter.Services.Strong.Core._2007_09.DataManagement.NamedReferenceInfo[1];

            refFileInfo[0] = new Teamcenter.Services.Strong.Core._2007_09.DataManagement.NamedReferenceInfo();

            refFileInfo[0].DeleteTarget = false;

            // refFileInfo[0].Type = "P7MCpkg";

            //refFileInfo[0].Type = "CD4CreoPackage";

            refFileInfo[0].Type = typeRef;

            //   refFileInfo[0].TargetObject = moDsRefList[0]; // To change


            /*** Named reference details ***/
            Teamcenter.Services.Strong.Core._2007_09.DataManagement.RemoveNamedReferenceFromDatasetInfo[] inputInfo =
                new Teamcenter.Services.Strong.Core._2007_09.DataManagement.RemoveNamedReferenceFromDatasetInfo[1];


            inputInfo[0] = new Teamcenter.Services.Strong.Core._2007_09.DataManagement.RemoveNamedReferenceFromDatasetInfo();

            inputInfo[0].ClientId = "DataRemove1";

            inputInfo[0].Dataset = (Dataset)dataSetMo;//.GetProperty("ref_list").ModelObjectValue;

            inputInfo[0].NrInfo = refFileInfo;


            /*** removing named reference from dataset***/

            // DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            ServiceData outputResponse = Tc_Services.dmService.RemoveNamedReferenceFromDataset(inputInfo);

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
        //==============================================================================================================
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
                log.write(logType.ERROR, "Exception in checkInModelObject in : " + e.ToString());

            }
            return isCheckedIn;

        }
        ///=================================================================================================
        public static void setIRProperty(ModelObject ModelObj, String value, String name)
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
                    log.write(logType.INFO,"setIRProperty Name : " + Prop_RealNameList[i]);
                    datasetPropery[0].VecNameVal[i].Name = Prop_RealNameList[i];
                    log.write(logType.INFO,"setIRProperty value : " + value);
                    datasetPropery[0].VecNameVal[i].Values = new String[] { value };
                }

                if (Tc_Services.dmService == null)
                {
                    log.write(logType.ERROR,"setIRProperty dmService is NULL: ");
                    return;
                }

                Teamcenter.Services.Strong.Core._2010_09.DataManagement.SetPropertyResponse resp = Tc_Services.dmService.SetProperties(datasetPropery, new String[0]);

                if (resp.Data.sizeOfPartialErrors() > 0)
                {
                    log.write(logType.ERROR,"setIRProperty partial Error count: " + resp.Data.sizeOfPartialErrors());
                }

                if (resp.Data.sizeOfUpdatedObjects() > 0)
                {
                    log.write(logType.INFO,"setIRProperty updated object count: " + resp.Data.sizeOfUpdatedObjects());
                }
            }
            catch (Exception ex)
            {
                log.write(logType.ERROR,"setIRProperty Exception: " + ex.Message);
                log.write(logType.ERROR, "setIRProperty Exception: " + ex.StackTrace);
                return;

            }

        }

        //public static void uploadExcelToTC_Test(string itemID, string RevID, string outputXLfileName)
        //{

           


        //    log.write(logType.INFO, "uploadExcelToTC: getItemRevisionQuery..");

        //    ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID);

        //    if (itemRevMO == null)
        //    {
        //        log.write(logType.ERROR, "uploadExcelToTC: item REV model Object is NULL/Empty");
        //        return;

        //    }

        //    log.write(logType.INFO, "uploadExcelToTC: isDataSetAvailable..");
        //    ModelObject dataSetMo = isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");

        //    if (dataSetMo == null)
        //    {
        //        log.write(logType.INFO, "uploadExcelToTC: createDataSet2..");
        //        Dataset dsCreated = createDataSet2(itemRevMO, itemID);
        //        log.write(logType.INFO, "uploadExcelToTC: attachFileToDataSet..");
        //        attachFileToDataSet(dsCreated, outputXLfileName, "MS ExcelX", itemID, "Test", "excel");
        //    }
        //    else
        //    {
        //        log.write(logType.INFO, "uploadExcelToTC: removeDataSetReference..");
        //        removeDataSetReference(dataSetMo, itemRevMO, "excel", "MS ExcelX");
        //        log.write(logType.INFO, "uploadExcelToTC: attachFileToDataSet..");
        //        attachFileToDataSet(dataSetMo, outputXLfileName, "MS ExcelX", itemID, "Test", "excel");
        //        log.write(logType.INFO, "uploadExcelToTC: checkInModelObject..");
        //        checkInModelObject(dataSetMo);
        //    }


        //}
        //============================================================================
    }
}
