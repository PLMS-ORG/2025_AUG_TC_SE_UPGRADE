using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Windows.Forms;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Common;
using Teamcenter.Soa.Client.Model.Strong;
using Teamcenter.Soa.Client;
using Teamcenter.Services.Strong.Query._2010_09.SavedQuery;
using System.IO;
using AddToTc.CDAT_BulkUploader;
using Log;
using DemoAddInTC.services;
using CreoToTc.Utils;


namespace Creo_TC_Live_Integration.TcDataManagement
{
    public static class DownloadDatasetNamedReference
    {
        public static String DownloadPropertiesTextFilePath = "";

        /*** Creating data management service and connection object ***/
        private static DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

        public static ModelObject getItemRev(String ItemID, String RevID)
        {
            // get services...
            Teamcenter.Services.Strong.Core.DataManagementService dmService = Teamcenter.Services.Strong.Core.DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());
            Teamcenter.Services.Strong.Query.SavedQueryService queryService = Teamcenter.Services.Strong.Query.SavedQueryService.getService(Teamcenter.ClientX.Session.getConnection());



            BusinessObjectQueryInput input = new BusinessObjectQueryInput();
            BusinessObjectQueryClause clause = new BusinessObjectQueryClause();
            //BusinessObjectQueryClause clause2 = new BusinessObjectQueryClause();

            clause.MathOperator = "=";
            clause.LogicOperator = "";
            clause.PropName = "item_id";
            clause.PropValue = ItemID;

            //clause2.MathOperator = "=";
            //clause2.LogicOperator = "AND";
            //clause2.PropName = "item_revision_id";
            //clause2.PropValue = RevID;

            input.ClientId = "";
            input.MaxNumToReturn = 0;
            input.TypeName = "Item";
            input.Clauses = new BusinessObjectQueryClause[] { clause };

            Teamcenter.Services.Strong.Query._2007_09.SavedQuery.SavedQueriesResponse newsavedQry_resp = queryService.ExecuteBusinessObjectQueries(new BusinessObjectQueryInput[] { input });
            // get ItemRevision ModelObject from query response
            ModelObject revModelObject = null;
            Teamcenter.Services.Strong.Query._2007_09.SavedQuery.QueryResults[] output_objs = newsavedQry_resp.ArrayOfResults;
            if (output_objs[0].ObjectUIDS.Length > 0)
            {
                String[] uid = output_objs[0].ObjectUIDS;
                ServiceData sData = dmService.LoadObjects(uid);
                revModelObject = (ModelObject)sData.GetPlainObject(0);
            }
            return revModelObject;
        }

        public static List<object> openBOMWindow2(ModelObject parentItemRevMO)
        {
            //  ModelObject[] itemRevisionModelObject = parentItemRevMO.GetProperty("revision_list").ModelObjectArrayValue;


            //  ItemRevision irObject = (ItemRevision)itemRevisionModelObject[0];
            Teamcenter.Soa.Client.Model.Strong.ItemRevision parentItemRev = (Teamcenter.Soa.Client.Model.Strong.ItemRevision)parentItemRevMO;

            List<object> bomWindowandParentLine = new List<object>(2);

            Teamcenter.Services.Strong.Cad._2007_01.StructureManagement.CreateBOMWindowsInfo[] createBOMWindowsInfo = new Teamcenter.Services.Strong.Cad._2007_01.StructureManagement.CreateBOMWindowsInfo[1];

            createBOMWindowsInfo[0] = new Teamcenter.Services.Strong.Cad._2007_01.StructureManagement.CreateBOMWindowsInfo();

            createBOMWindowsInfo[0].ItemRev = (Teamcenter.Soa.Client.Model.Strong.ItemRevision)parentItemRev;

            Teamcenter.Services.Strong.Cad._2007_01.StructureManagement.StructureManagement sms2 = Teamcenter.Services.Strong.Cad.StructureManagementService.getService(Teamcenter.ClientX.Session.getConnection());
            //   Teamcenter.Services.Strong.Cad.StructureManagementService sms2 = Teamcenter.Services.Strong.Cad.StructureManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            Teamcenter.Services.Strong.Cad._2007_01.StructureManagement.CreateBOMWindowsResponse createBOMWindowsResponse = sms2.CreateBOMWindows(createBOMWindowsInfo);



            if (createBOMWindowsResponse.ServiceData.sizeOfPartialErrors() > 0)
            {
                for (int i = 0; i < createBOMWindowsResponse.ServiceData.sizeOfPartialErrors(); i++)
                {

                    //System.out.println("Partial Error in Open BOMWindow = "+createBOMWindowsResponse.serviceData.getPartialError(i).getMessages()[0]);
                }
            }

            bomWindowandParentLine.Add(createBOMWindowsResponse.Output[0].BomWindow);//BOMWindow

            bomWindowandParentLine.Add(createBOMWindowsResponse.Output[0].BomLine);//TOPLine in BOMWINDOW
            return bomWindowandParentLine;
        }

        //public static void getBomStructure(string itemId, String revId, String StageDir, bool parentRevDataSetDownload, ModelObject itemRevObj)
        //{


        //    String properties = "";// ======>**************


        //    ModelObject parentItem = itemRevObj;
        //    String Item = itemId;
        //    List<Object> bomWindowAndLine = openBOMWindow2(parentItem);//open BOM WINDOW
        //    Teamcenter.Soa.Client.Model.Strong.BOMWindow bomWindow = (Teamcenter.Soa.Client.Model.Strong.BOMWindow)bomWindowAndLine[0];
        //    Teamcenter.Soa.Client.Model.Strong.BOMLine parentBomLine = (Teamcenter.Soa.Client.Model.Strong.BOMLine)bomWindowAndLine[1];
        //    Teamcenter.Services.Strong.Core.DataManagementService dmService = Teamcenter.Services.Strong.Core.DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

        //    dmService.GetProperties(new ModelObject[] { parentBomLine }, new String[] { "bl_rev_ps_children" });

        //    string listOfChildren = parentBomLine.Bl_rev_ps_children;
        //    bomWindowAndLine.Clear();
        //    bomWindow = null;
        //    if (parentRevDataSetDownload == true)
        //    {
        //        // Download EndItemRevisions DataSet
        //        ModelObject itemRevObject = DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemId, revId, StageDir, true,
        //            parentRevDataSetDownload);
        //    }
        //    char delimiter = ',';
        //    List<String> childrenItems = null;
        //    // childrenItems = (DataProcess.stringSplit(listOfChildren, delimiter));
        //    childrenItems = (stringSplitForDownloadFunction(listOfChildren, delimiter));



        //    if (listOfChildren.CompareTo("") != 0)
        //    {
        //        DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemId, revId, StageDir, true, logFilePath, parentRevDataSetDownload);
        //        foreach (String childItemID in childrenItems)
        //        {

        //            String[] tokens = childItemID.Split('/');
        //            ModelObject itemRevObject = DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(tokens[0], tokens[1], StageDir,
        //                false, logFilePath, parentRevDataSetDownload);

        //            // recursive call of same method
        //            getBomStructure(tokens[0], tokens[1], StageDir, false, itemRevObject, logFilePath);

        //        }

        //    }
        //    else
        //    {

        //        DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemId, revId, StageDir, true, logFilePath
        //            , parentRevDataSetDownload);
        //    }

        //}

        public static List<String> stringSplitForDownloadFunction(string str, char delimiter, String logFilePath)
        {
            List<string> returnStringList = new List<String>();

            String[] returnStringArray = null;

            // Utlity.Log("To Split String :" + str + " || Delimiter :" + delimiter, logFilePath);
            returnStringArray = str.Split(delimiter);

            foreach (String temp in returnStringArray)
            {

                String[] tempString = null;
                tempString = temp.Split('-');
                returnStringList.Add(tempString[0].Replace(" ", ""));

            }
            return returnStringList;
        }
 
        public static void setObjectPolicy()
        {

            String[] props = new String[8];

            //   Creo_TC_Live_Integration.TeamCenter.TcData.tc_ItemRev_PropertynameHolder.CopyTo(props, 0);

            props[props.Count() - 8] = "cd4InstanceIdCp";

            props[props.Count() - 7] = "current_revision_id";

            props[props.Count() - 6] = "checked_out";

            props[props.Count() - 5] = "isObsolete";

            props[props.Count() - 4] = "revision_list";

            props[props.Count() - 3] = "item_id";

            props[props.Count() - 2] = "item_revision_id";

            props[props.Count() - 1] = "IMAN_specification";

            SessionService session = SessionService.getService(Teamcenter.ClientX.Session.getConnection());

            ObjectPropertyPolicy policy = new ObjectPropertyPolicy();

            //  policy.AddType(new PolicyType("ItemRevision", props));
            //prepare Input details to set Objectpoicy on given Item and for given property
            Teamcenter.Soa.Common.ObjectPropertyPolicy objPropertyPolicy = new Teamcenter.Soa.Common.ObjectPropertyPolicy();
            Teamcenter.Soa.Common.PolicyType policyType = new Teamcenter.Soa.Common.PolicyType("BOMLine", new String[] { "bl_item_item_id", "ps_children", "Bl_child_lines", "bl_all_notes", "bl_bomview_rev", "bl_bomview", "CD4InstanceId", "cd4InstanceId", "bl_line_name", "bl_sequence_no", "bl_plmxml_abs_xform", "bl_all_notes", "bl_quantity", "SE ObjectID", "ps_children", "bl_rev_ps_children", "bl_all_child_lines", "bl_child_lines" });
            Teamcenter.Soa.Common.PolicyType policyTypeItem = new Teamcenter.Soa.Common.PolicyType("Item", new String[] { "object_name", "revision_list" });
            Teamcenter.Soa.Common.PolicyType policyTypeItemRev = new Teamcenter.Soa.Common.PolicyType("ItemRevision", props);

            Teamcenter.Soa.Common.PolicyType policyDs = new Teamcenter.Soa.Common.PolicyType("Dataset", new String[] { "type_name", "dataset", "original_file_name", "object_string", "cd4UploadTime", "object_desc", "ref_list", "checked_out", "dataset_type", "datasettype_name", "object_type" });

            Teamcenter.Soa.Common.PolicyType policyNR = new Teamcenter.Soa.Common.PolicyType("ImanFile", new String[] { "original_file_name" });

            objPropertyPolicy.AddType(new PolicyType("DatasetType", new String[] { "dataset", "ref_list", "original_file_name", "cd4UploadTime", "Dataset_type", "datasettype_name", "type_name" }));

            //  policy.AddType(new PolicyType("Dataset", new string[] { "ref_list" }));
            objPropertyPolicy.AddType(policyDs);
            objPropertyPolicy.AddType(policyNR);
            objPropertyPolicy.AddType(policyType);
            objPropertyPolicy.AddType(policyTypeItem);
            objPropertyPolicy.AddType(policyTypeItemRev);
            //get ObjectPolicyService
            Teamcenter.Services.Strong.Core.SessionService ss = Teamcenter.Services.Strong.Core.SessionService.getService(Teamcenter.ClientX.Session.getConnection());
            //set ObjectPolicy
            ss.SetObjectPropertyPolicy(objPropertyPolicy);
        }

        //=====================================================================================================
        public static ModelObject RetrieveItemRevMOAndDownloadDatasetNR(String ItemID, String RevID, String StageDir, Boolean dataSetDownloadFlag,
                                                                        Boolean parentItemRevDownloadFlag)
        {
            try
            {
                //TcAdaptor.setObjectPolicy_General(logFilePath);

                ModelObject revModelObj = null;
                revModelObj = getItemRevisionQuery(ItemID, RevID);
                if (revModelObj == null)
                {
                    log.write(logType.ERROR, "Item Rev Model Object is NULL for: " + ItemID + RevID);
                    return null;
                }
                log.write(logType.INFO,"RetrieveItemRevMOAndDownloadDatasetNR: ItemID: " + ItemID);
                log.write(logType.INFO, "RetrieveItemRevMOAndDownloadDatasetNR: RevID: " + RevID);

                //DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

                //dmService.RefreshObjects(new ModelObject[] { revModelObj });
               // dmService.LoadObjects(new string[] { revModelObj.Uid });



                if (dataSetDownloadFlag)
                {
                    ModelObject[] dataSets = getAllDataSet(revModelObj);
                    if (dataSets != null)
                    {

                        foreach (ModelObject dsMo in dataSets)
                        {

                           // ModelObject[] mos = new ModelObject[] { dsMo };
                            //dmService.RefreshObjects(mos);

                            Dataset ds = (Dataset)dsMo;
                            String dsUID = ds.Uid;
                            String[] arr = new String[] { dsUID };
                            ServiceData sData = dmService.LoadObjects(arr);
                            Dataset dsMObj = (Dataset)sData.GetPlainObject(0);

                            ServiceData sData2 = Tc_Services.dmService.LoadObjects(arr);

                            Tc_Services.dmService.RefreshObjects(new ModelObject[] { ds });


                            String object_Type = dsMObj.Object_type;
                            log.write(logType.INFO, "DsType:" + object_Type);

                            if (object_Type.Equals("MSExcelX", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                //TcAdaptor.checkOutModelObject(dsMo);
                                getNamedReferenceFile(ds, StageDir, parentItemRevDownloadFlag);
                              //  getNamedReferenceFile(dsMObj, StageDir, logFilePath, parentItemRevDownloadFlag);
                            }
                        }

                    }
                    else
                    {
                        log.write(logType.ERROR, "DataSet is empty for ItemID/RevID :" + ItemID + "/" + RevID   );
                    }
                }

                return revModelObj;
            }

            catch (Exception e)
            {

                log.write(logType.ERROR, "Exception in RetrieveItemRevMOAndDownloadDatasetNR function :" + e.Message + ". ItemID/RevID" + ItemID + "/" + RevID);
            }

            return null;
        }
        //=====================================================================================================

        //20-01-20 Methun
        public static bool checkForExcelDataset(String ItemID, String RevID, String StageDir, Boolean dataSetDownloadFlag, Boolean parentItemRevDownloadFlag)
        {
            try
            {
                //setObjectPolicy();
                ModelObject revModelObj = null;
                log.write(logType.ERROR, "checkForExcelDataset: ItemID: " + ItemID);
                log.write(logType.ERROR, "checkForExcelDataset: RevID: " + RevID);
                revModelObj = getItemRevisionQuery(ItemID, RevID);

                if (revModelObj == null)
                {

                    log.write(logType.ERROR, "checkForExcelDataset: Item Rev Model Object is NULL: " + ItemID);
                    return false;
                }


                ModelObject[] dataSets = getAllDataSet(revModelObj);
                if (dataSets != null)
                {

                    foreach (ModelObject dsMo in dataSets)
                    {
                        Dataset ds = (Dataset)dsMo;

                        String dsUID = ds.Uid;

                        DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

                        String[] arr = new String[] { dsUID };

                        ServiceData sData = dmService.LoadObjects(arr);

                        Dataset dsMObj = (Dataset)sData.GetPlainObject(0);
                        String object_Type = dsMObj.Object_type;
                        log.write(logType.ERROR, "checkForExcelDataset:  DsType:" + object_Type);

                        if (object_Type.Equals("MSExcelX", StringComparison.OrdinalIgnoreCase) == true)
                            return true;
                    }

                }
                else
                {
                    log.write(logType.ERROR,"checkForExcelDataset: DataSet is empty for ItemID/RevID :" + ItemID + "/" + RevID);
                    return false;
                }
            }

            catch (Exception e)
            {

                log.write(logType.ERROR, "checkForExcelDataset: Exception in function :" + e.Message + ". ItemID/RevID" + ItemID + "/" + RevID);
                return false;
            }

            return false;
        }
        //20-01-20 Methun

        public static ModelObject[] getAllDataSet(ModelObject itemRevisionModelObject)
        {
            try
            {
                //TcAdaptor.setObjectPolicy_General(logFilePath);
                ItemRevision itemRev = (ItemRevision)itemRevisionModelObject;
                if (itemRev == null)
                {
                    log.write(logType.ERROR, "isDataSetAvailable: " + "itemRev is null...");
                    return null;
                }

                ModelObject[] datasetMoArray = itemRev.IMAN_specification;
                //  String based_on = itemRev.Based_on;

                if (datasetMoArray != null && datasetMoArray.Length > 0)
                {
                    return datasetMoArray;
                }
                return null;

            }

            catch (Exception e)
            {
                log.write(logType.ERROR, "exception in getAllDataSet function" + e.Message);
            }

            return null;
        }


        public static ImanQuery getTcQuery(String queryToFind)
        {

            ImanQuery tcQueryToReturn = null;

            Teamcenter.Services.Strong.Query.SavedQueryService savedQryServices = Teamcenter.Services.Strong.Query.SavedQueryService.getService(Teamcenter.ClientX.Session.getConnection());

            Teamcenter.Services.Strong.Query._2006_03.SavedQuery.GetSavedQueriesResponse savedQueries = savedQryServices.GetSavedQueries();

            //Utlity.Log("Searching saved queries in Teamcenter", logFilePath);

            if (savedQueries.Queries.Length == 0)
            {

                log.write(logType.ERROR,"Failed to get saved queries");
                return null;
            }

            else
            {
                for (int i = 0; i < savedQueries.Queries.Length; i++)
                {
                    if (savedQueries.Queries[i].Name.Equals(queryToFind))
                    {
                        tcQueryToReturn = savedQueries.Queries[i].Query;

                        //Utlity.Log("Identified ItemRevision... saved query in Teamcenter", logFilePath);

                        break;
                    }
                }
            }
            return tcQueryToReturn;
        }

        public static ModelObject getItemRevisionQuery(String ItemID, String RevisionID)
        {
            try
            {
                ImanQuery qry = null;

                qry = getTcQuery("Item Revision...");

                Teamcenter.Services.Strong.Query.SavedQueryService savedQryServices = Teamcenter.Services.Strong.Query.SavedQueryService.getService(Teamcenter.ClientX.Session.getConnection());

                if (savedQryServices == null)
                {
                    log.write(logType.ERROR, "getItemRevisionQuery: " + "savedQryServices is NULL: ");
                    return null;
                }

                Teamcenter.Services.Strong.Query._2006_03.SavedQuery.GetSavedQueriesResponse savedQueries = savedQryServices.GetSavedQueries();


                if (savedQueries == null)
                {
                    log.write(logType.ERROR, "getItemRevisionQuery: " + "savedQueries is NULL: ");
                    return null;
                }

                ModelObject itemRevModelObject = null;
                /*** Finding itemrevision by sysid using saved query ***/

                if (qry != null)
                {
                    //Utlity.Log("Inside getItemRevisionQuery", logFilePath);

                    //setObjectPolicy();

                    //Teamcenter.Services.Strong.Core.DataManagementService dmService = Teamcenter.Services.Strong.Core.DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());



                    Teamcenter.Services.Strong.Query._2007_06.SavedQuery.SavedQueryInput[] qryInput = new Teamcenter.Services.Strong.Query._2007_06.SavedQuery.SavedQueryInput[1];


                    qryInput[0] = new Teamcenter.Services.Strong.Query._2007_06.SavedQuery.SavedQueryInput();

                    qryInput[0].Query = qry;

                    qryInput[0].MaxNumToReturn = 0; // 0 means no limit

                    qryInput[0].Entries = new String[] { "Item ID", "Revision" };

                    qryInput[0].Values = new String[2];
                    ItemID = ItemID.Trim();
                    RevisionID = RevisionID.Trim();
                    qryInput[0].Values[0] = ItemID;
                    qryInput[0].Values[1] = RevisionID;



                    Teamcenter.Services.Strong.Query._2007_06.SavedQuery.ExecuteSavedQueriesResponse executeQry = savedQryServices.ExecuteSavedQueries(qryInput);

                    if (executeQry == null)
                    {
                        log.write(logType.ERROR, "getItemRevisionQuery: " + "executeQry is NULL: ");
                        return null;
                    }

                    Teamcenter.Services.Strong.Query._2007_06.SavedQuery.SavedQueryResults qryResult = executeQry.ArrayOfResults[0];
                    if (qryResult == null)
                    {
                        log.write(logType.ERROR, "getItemRevisionQuery: " + "qryResult is NULL: ");
                        return null;
                    }


                    // Assumption - Last Entry in this Query Result is the latest Revision... (Doubtful ??)
                    if (qryResult.Objects.Length > 0)
                    {

                        //setObjectPolicy();

                        String uid = qryResult.Objects[qryResult.Objects.Count() - 1].Uid;

                        if (dmService == null)
                        {
                            log.write(logType.ERROR, "getItemRevisionQuery: " + "dmService is NULL: ");
                            return null;
                        }

                        ServiceData sData = dmService.LoadObjects(new String[] { uid });

                        itemRevModelObject = (ModelObject)sData.GetPlainObject(0);

                        dmService.RefreshObjects(new ModelObject[] { itemRevModelObject });

                        //Utlity.Log("Queryresult Count/Length ::" + qryResult.Objects.Count() + "/" + qryResult.Objects.Length, logFilePath);


                    }
                }

                else
                {
                    //Utlity.Log("getItemRevisionQuery: " + "Item Revisions... query not found in Teamcenter.", logFilePath);
                }

                //Utlity.Log( "Checking Item Revision in Teamcenter by Item ID & Rev ID process is completed.",logFilePath);


                return itemRevModelObject;
            }
            catch (Exception e)
            {
                log.write(logType.ERROR, "getItemRevisionQuery: " + "Exception in CheckItemIdBySysID functions: " + e.Message);
                log.write(logType.ERROR, "getItemRevisionQuery: " + "Exception in CheckItemIdBySysID functions: " + e.StackTrace);

            }
            return null;
        }
        public static void getNamedReferenceFile(Dataset ds, String StageDir,
            Boolean parentItemRevDownloadFlag)
        {
            try
            {
                if (ds == null)
                {
                    log.write(logType.ERROR, "getNamedReferenceFile: " + "ds is" + " Empty");
                    return;
                }
                log.write(logType.INFO, "getNamedReferenceFile: StageDir:" + StageDir);


                //String fscHost = TCPropertyReader.get(Constants.TC_FSC_HOST);
                //if (fscHost == null || fscHost.Equals("") == true)
                //{
                //    log.write(logType.ERROR, "getNamedReferenceFile: " + "fscHost is" + " Empty");
                //    return;

                //}
                //String[] url = { fscHost };

                if (Teamcenter.ClientX.Session.getConnection() == null)
                {
                    log.write(logType.ERROR, "getNamedReferenceFile:" + "No Teamcenter Connection...");
                    return;
                }

                //String serverHost = TCPropertyReader.get(Constants.TC_SERVER_HOST);
                ////String serverHost = Constants.TC_SERVER_HOST;
                //if (serverHost == null || serverHost.Equals("") == true)
                //{
                //    log.write(logType.ERROR, "getNamedReferenceFile: " + "serverHost is" + " Empty");
                //    return;

                //}

                // FileManagementUtility fileUtility = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection(), serverHost, url, url, StageDir);
                DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

                if (Tc_Services.fileManag_Utility == null)
                {
                    log.write(logType.ERROR, "getNamedReferenceFile: " + "file Management Utility is" + " Empty");
                    return;
                }

                if (dmService == null)
                {
                    log.write(logType.ERROR, "getNamedReferenceFile: " + "dmService is" + " Empty");
                    return;
                }

                if (ds == null)
                {
                    log.write(logType.ERROR, "getNamedReferenceFile: " + "dataset is" + " Empty");
                    return;
                }

                //Teamcenter.Soa.Client.FileManagementUtility fileMangeUtil = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection(), "139.64.32.153", url, url, StageDir);
                ModelObject[] NRfiles = ds.Ref_list;
                log.write(logType.INFO, "NRfiles count: " + NRfiles.Length);
                StringBuilder sb = new StringBuilder();
                //for (int jdx = 0; jdx < NRfiles.Length; jdx++)
                foreach (ModelObject NRfile in NRfiles)
                {

                    // Load NR file and get its original name
                    String NRUid = NRfile.Uid;
                    String[] NRUidArray = new String[] { NRUid };
                    dmService.LoadObjects(NRUidArray);
                    String NRname = NRfile.GetPropertyDisplayableValue("original_file_name");
                    //  String name = NRfile.GetProperty("File_name").ToString();
                    // get NR file to the given StageDir => fullFilePath along with fileName
                    if (!NRname.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) == true) continue;
                    //if (NRname.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) == true) continue;
                    //if (NRname.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase) == true) continue;
                    //if (NRname.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase) == true) continue;
                    String DestinationFileFullPath = System.IO.Path.Combine(StageDir, NRname);
                    log.write(logType.INFO, "File is: " + DestinationFileFullPath);
                    if (File.Exists(DestinationFileFullPath) == true)
                    {
                        File.Delete(DestinationFileFullPath);
                    }
                    Object tempObj = new Object();                 
                    if (Tc_Services.fileManag_Utility == null)
                        log.write(logType.ERROR, "fileUtility1 is null");
                    log.write(logType.INFO, "NRUid: " + NRUid);

                    GetFileResponse res = Tc_Services.fileManag_Utility.GetFileToLocation(NRfile, DestinationFileFullPath, null, tempObj);
                    log.write(logType.INFO, "GetFileToLocation is completed" );

                    System.IO.FileInfo[] files = res.GetFiles();
                    log.write(logType.INFO, "fileName: " + files[0].FullName);

                    for (int j = 0; j < files.Length; j++)
                    {
                        String path = files[j].FullName;
                        log.write(logType.INFO, "filePath: " + path);
                        if (parentItemRevDownloadFlag == true)
                        {
                            sb.AppendLine(path);
                        }
                    }

                    if (parentItemRevDownloadFlag == true)
                    {
                        if (sb != null && sb.Length != 0)
                        {
                            // Signal to Creo to Open this FILE......
                            CreoUtilitySession.SignalToCreoToOpenFileDownloadedFromTC = sb.ToString();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                log.write(logType.ERROR, "Exception @ getNamedReferenceFile::" + e.Message);
                log.write(logType.ERROR, "Exception @ getNamedReferenceFile::" + e.StackTrace);
                log.write(logType.ERROR, "Exception @ getNamedReferenceFile::" + e.InnerException);

            }
        }


        public static void WriteToDownloadText(string Info, string DownloadTextFilePath, String logFilePath)
        {
            try
            {
                StreamWriter w = File.AppendText(DownloadTextFilePath);
                w.WriteLine("{0}", Info);
                Console.WriteLine(Info);
                w.Close();
            }
            catch (Exception ex)
            {
                log.write(logType.ERROR, "WriteToDownloadText: Log Writing Exception: " + ex.Message);
            }

        }

        //public static bool check_BvrIs_AlreadyInTc(ModelObject revisionModelObj, String logFilePath)
        //{

        //    try
        //    {
        //        //setObjectPolicy();
        //        List<Object> bomWindowAndLine = CreateBom.open_BomWindow_From_ItemRevMo((ItemRevision)revisionModelObj, logFilePath);//open BOM WINDOW
        //        if (bomWindowAndLine == null || bomWindowAndLine.Count == 0)
        //        {
        //            Utlity.Log("WARNING :check_BvrIs_AlreadyInTc returned False", logFilePath);
        //            return false;
        //        }

        //        Teamcenter.Soa.Client.Model.Strong.BOMWindow bomWindow = (Teamcenter.Soa.Client.Model.Strong.BOMWindow)bomWindowAndLine[0];

        //        Teamcenter.Soa.Client.Model.Strong.BOMLine parentBomLine = (Teamcenter.Soa.Client.Model.Strong.BOMLine)bomWindowAndLine[1];

        //        dmService.GetProperties(new ModelObject[] { parentBomLine }, new String[] { "bl_rev_ps_children" });

        //        string listOfChildren = parentBomLine.Bl_rev_ps_children;

        //        ModelObject[] childModelObject = parentBomLine.Bl_child_lines;

        //        if (childModelObject.Count() == 0)
        //        {

        //            Utlity.Log("WARNING :Top item does not have bvr", logFilePath);

        //            Console.WriteLine("WARNING :Top item does not have bvr");

        //            return false;
        //        }

        //        return true;
        //    }

        //    catch (Exception e)
        //    {
        //        //Utlity.Log("ERROR :Exception in check_BvrIs_AlreadyInTc function ." + e.Message, logFilePath);

        //        //Utlity.Log("ERROR :Exception in check_BvrIs_AlreadyInTc function ." + e.StackTrace, logFilePath);

        //        Console.WriteLine("ERROR :Exception in check_BvrIs_AlreadyInTc function ." + e.StackTrace);

        //        return false;
        //    }

        //}

        //=====================================================================================
    }
}
