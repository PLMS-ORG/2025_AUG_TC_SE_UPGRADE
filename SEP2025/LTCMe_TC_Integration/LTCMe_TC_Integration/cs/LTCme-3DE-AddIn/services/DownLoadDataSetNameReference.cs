using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Services.Strong.Query;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Common;
using Teamcenter.Soa.Client.Model.Strong;
using Teamcenter.Services.Strong.Core._2008_06.DataManagement;
using Teamcenter.Soa.Client;
using Teamcenter.ClientX;
using Teamcenter.Services.Strong.Query._2010_09.SavedQuery;
using System.IO;
using DemoAddInTC.services;
using DemoAddInTC.utils;
using DemoAddInTC;
using DemoAddInTC.se;


namespace Creo_TC_Live_Integration.TcDataManagement
{
    public static class DownloadDatasetNamedReference
    {
        public static String DownloadPropertiesTextFilePath = "";

        /*** Creating data management service and connection object ***/
        private static DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

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

        public static void getBomStructure(string itemId, String revId, String StageDir, bool parentRevDataSetDownload, ModelObject itemRevObj,
            String logFilePath)
        {


            String properties = "";// ======>**************


            ModelObject parentItem = itemRevObj;
            String Item = itemId;
            List<Object> bomWindowAndLine = openBOMWindow2(parentItem);//open BOM WINDOW
            Teamcenter.Soa.Client.Model.Strong.BOMWindow bomWindow = (Teamcenter.Soa.Client.Model.Strong.BOMWindow)bomWindowAndLine[0];
            Teamcenter.Soa.Client.Model.Strong.BOMLine parentBomLine = (Teamcenter.Soa.Client.Model.Strong.BOMLine)bomWindowAndLine[1];
            Teamcenter.Services.Strong.Core.DataManagementService dmService = Teamcenter.Services.Strong.Core.DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

            dmService.GetProperties(new ModelObject[] { parentBomLine }, new String[] { "bl_rev_ps_children" });

            string listOfChildren = parentBomLine.Bl_rev_ps_children;
            bomWindowAndLine.Clear();
            bomWindow = null;
            if (parentRevDataSetDownload == true)
            {
                // Download EndItemRevisions DataSet
                ModelObject itemRevObject = DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemId, revId, StageDir, true, logFilePath,
                    parentRevDataSetDownload);
            }
            char delimiter = ',';
            List<String> childrenItems = null;
            // childrenItems = (DataProcess.stringSplit(listOfChildren, delimiter));
            childrenItems = (stringSplitForDownloadFunction(listOfChildren, delimiter, logFilePath));



            if (listOfChildren.CompareTo("") != 0)
            {
                DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemId, revId, StageDir, true, logFilePath, parentRevDataSetDownload);
                foreach (String childItemID in childrenItems)
                {

                    String[] tokens = childItemID.Split('/');
                    ModelObject itemRevObject = DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(tokens[0], tokens[1], StageDir,
                        false, logFilePath, parentRevDataSetDownload);

                    // recursive call of same method
                    getBomStructure(tokens[0], tokens[1], StageDir, false, itemRevObject, logFilePath);

                }

            }
            else
            {

                DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(itemId, revId, StageDir, true, logFilePath
                    , parentRevDataSetDownload);
            }

        }

        public static List<String> stringSplitForDownloadFunction(string str, char delimiter, String logFilePath)
        {
            List<string> returnStringList = new List<String>();

            String[] returnStringArray = null;

            // Utility.Log("To Split String :" + str + " || Delimiter :" + delimiter, logFilePath);
            returnStringArray = str.Split(delimiter);

            foreach (String temp in returnStringArray)
            {

                String[] tempString = null;
                tempString = temp.Split('-');
                returnStringList.Add(tempString[0].Replace(" ", ""));

            }
            return returnStringList;
        }

        /*      private static string getProperties(ModelObject itemRevObj)
              {

                  StringBuilder sb = new StringBuilder();
                  String contentID = itemRevObj.GetPropertyDisplayableValue("cd4InstanceIdCp");
                  sb.Append(contentID);
                  sb.Append(Constants.TILDE);
                  String itemID = itemRevObj.GetPropertyDisplayableValue("item_id");
                  sb.Append(itemID);
                  sb.Append(Constants.TILDE);
                  String revID = itemRevObj.GetPropertyDisplayableValue("item_revision_id");
                  sb.Append(revID);
                  sb.Append(Constants.TILDE);

                  foreach (String propertyname in Creo_TC_Live_Integration.TeamCenter.TcData.tc_ItemRev_PropertynameHolder)
                  {
                      if (Creo_TC_Live_Integration.TeamCenter.TcData.mapUiHeaderMapper.ContainsKey(propertyname))
                      {
                          String displayName = "";

                          Creo_TC_Live_Integration.TeamCenter.TcData.mapUiHeaderMapper.TryGetValue(propertyname, out displayName);

                          String mValue = itemRevObj.GetPropertyDisplayableValue(propertyname);

                          sb.Append(mValue); ;

                          sb.Append(Constants.TILDE);

                          Teamcenter.Soa.Client.Model.Property p = (Teamcenter.Soa.Client.Model.Property)itemRevObj.GetProperty(propertyname);

                          //  String s = p.StringValue;

                          //  System.Collections.IList l=p.ModelObjectListValue;
                          //    // Utility.Log("Props Values :"+s);

                          //   co.updateCreoProperties(displayName, mValue);
                      }

                      else
                      {
                          //    co.updateCreoProperties(propertyname, latestRevMo.GetPropertyDisplayableValue(propertyname));

                          //   // Utility.Log("WARNING : Teamcenter real name showing in UI header .Please check your header mapper configuration file for display name for property ->" + propertyname, logFilePath);
                      }

                  }
                  String properties = sb.ToString();

                  properties = properties.TrimEnd('~');

                  return properties;
              }
              */
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


        public static ModelObject RetrieveItemRevMOAndDownloadDatasetNR(String ItemID, String RevID, String StageDir, Boolean dataSetDownloadFlag, String logFilePath, Boolean parentItemRevDownloadFlag)
        {
            try
            {
                //setObjectPolicy();
                Utlity.Log("setObjectPolicy_General Start..", logFilePath);
                TcAdaptor.setObjectPolicy_General(logFilePath);
                Utlity.Log("setObjectPolicy_General End..", logFilePath);
                ModelObject revModelObj = null;
                Utlity.Log("dmService Start..", logFilePath);
                DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());
                Utlity.Log("dmService End..", logFilePath);
                Utility.Log("RetrieveItemRevMOAndDownloadDatasetNR: ItemID: " + ItemID, logFilePath);
                Utility.Log("RetrieveItemRevMOAndDownloadDatasetNR: RevID: " + RevID, logFilePath);
                revModelObj = getItemRevisionQuery(ItemID, RevID, logFilePath);
                if (revModelObj == null)
                {

                    Utility.Log("Item Rev Model Object is NULL: " + ItemID, logFilePath);
                    return null;
                }

                dmService.RefreshObjects(new ModelObject[] {revModelObj});
                dmService.LoadObjects(new string[] { revModelObj.Uid });

                

                if (dataSetDownloadFlag)
                {
                    ModelObject[] dataSets = getAllDataSet(revModelObj, logFilePath);
                    if (dataSets != null)
                    {

                        foreach (ModelObject dsMo in dataSets)
                        {
                            
                            ModelObject[] mos = new ModelObject[] { dsMo };
                            dmService.RefreshObjects(mos);
                            
                            Dataset ds = (Dataset)dsMo;

                            String dsUID = ds.Uid;
                            
                            //String object_type = dsMo.GetPropertyDisplayableValue("object_type");

                            String[] arr = new String[] { dsUID };
                            

                            ServiceData sData = dmService.LoadObjects(arr);

                            Dataset dsMObj = (Dataset)sData.GetPlainObject(0);
                            String object_Type = dsMObj.Object_type;
                            Utility.Log("DsType:" + object_Type, logFilePath);

                            if (object_Type.Equals("MSExcelX", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                //TcAdaptor.checkOutModelObject(dsMo);
                                getNamedReferenceFile(dsMObj, StageDir, logFilePath, parentItemRevDownloadFlag);
                            }
                        }

                    }
                    else
                    {
                        Utility.Log("DataSet is empty for ItemID/RevID :" + ItemID + "/" + RevID, logFilePath);
                    }
                }

                return revModelObj;
            }

            catch (Exception e)
            {

                Utility.Log("Exception in RetrieveItemRevMOAndDownloadDatasetNR function :" + e.Message + ". ItemID/RevID" + ItemID + "/" + RevID, logFilePath);
            }

            return null;
        }


        //20-01-20 Methun
        public static bool checkForExcelDataset(String ItemID, String RevID, String StageDir, Boolean dataSetDownloadFlag, String logFilePath, Boolean parentItemRevDownloadFlag)
        {
            try
            {
                //setObjectPolicy();
                ModelObject revModelObj = null;
                Utility.Log("checkForExcelDataset: ItemID: " + ItemID, logFilePath);
                Utility.Log("checkForExcelDataset: RevID: " + RevID, logFilePath);
                revModelObj = getItemRevisionQuery(ItemID, RevID, logFilePath);

                if (revModelObj == null)
                {

                    Utility.Log("checkForExcelDataset: Item Rev Model Object is NULL: " + ItemID, logFilePath);
                    return false;
                }


                ModelObject[] dataSets = getAllDataSet(revModelObj, logFilePath);
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
                        Utility.Log("checkForExcelDataset:  DsType:" + object_Type, logFilePath);

                        if (object_Type.Equals("MSExcelX", StringComparison.OrdinalIgnoreCase) == true)
                            return true;
                    }

                }
                else
                {
                    Utility.Log("checkForExcelDataset: DataSet is empty for ItemID/RevID :" + ItemID + "/" + RevID, logFilePath);
                    return false;
                }
            }

            catch (Exception e)
            {

                Utility.Log("checkForExcelDataset: Exception in function :" + e.Message + ". ItemID/RevID" + ItemID + "/" + RevID, logFilePath);
                return false;
            }

            return false;
        }
        //20-01-20 Methun

        public static ModelObject[] getAllDataSet(ModelObject itemRevisionModelObject, String logFilePath)
        {
            try
            {

                ItemRevision itemRev = (ItemRevision)itemRevisionModelObject;
                if (itemRev == null)
                {
                    Utility.Log("isDataSetAvailable: " + "itemRev is null...", logFilePath);
                    return null;
                }

                ModelObject[] datasetMoArray = itemRev.IMAN_specification;
                String based_on = itemRev.Based_on;

                if (datasetMoArray != null && datasetMoArray.Length > 0)
                {
                    return datasetMoArray;
                }
                return null;

            }

            catch (Exception e)
            {
                Utility.Log("exception in getAllDataSet function" + e.Message, logFilePath);
            }

            return null;
        }


        public static ImanQuery getTcQuery(String queryToFind, String logFilePath)
        {

            ImanQuery tcQueryToReturn = null;

            Teamcenter.Services.Strong.Query.SavedQueryService savedQryServices = Teamcenter.Services.Strong.Query.SavedQueryService.getService(Teamcenter.ClientX.Session.getConnection());

            Teamcenter.Services.Strong.Query._2006_03.SavedQuery.GetSavedQueriesResponse savedQueries = savedQryServices.GetSavedQueries();

            //Utility.Log("Searching saved queries in Teamcenter", logFilePath);

            if (savedQueries.Queries.Length == 0)
            {

                Utility.Log("Failed to get saved queries", logFilePath);
                return null;
            }

            else
            {
                for (int i = 0; i < savedQueries.Queries.Length; i++)
                {
                    if (savedQueries.Queries[i].Name.Equals(queryToFind))
                    {
                        tcQueryToReturn = savedQueries.Queries[i].Query;

                        //Utility.Log("Identified ItemRevision... saved query in Teamcenter", logFilePath);

                        break;
                    }
                }
            }
            return tcQueryToReturn;
        }

        public static ModelObject getItemRevisionQuery(String ItemID, String RevisionID, String logFilePath)
        {
            try
            {

                ImanQuery qry = null;

                qry = getTcQuery("Item Revision...", logFilePath);

                Teamcenter.Services.Strong.Query.SavedQueryService savedQryServices = Teamcenter.Services.Strong.Query.SavedQueryService.getService(Teamcenter.ClientX.Session.getConnection());

                if (savedQryServices == null)
                {
                    Utility.Log("getItemRevisionQuery: " + "savedQryServices is NULL: ", logFilePath);
                    return null;
                }

                Teamcenter.Services.Strong.Query._2006_03.SavedQuery.GetSavedQueriesResponse savedQueries = savedQryServices.GetSavedQueries();


                if (savedQueries == null)
                {
                    Utility.Log("getItemRevisionQuery: " + "savedQueries is NULL: ", logFilePath);
                    return null;
                }

                ModelObject itemRevModelObject = null;
                /*** Finding itemrevision by sysid using saved query ***/

                if (qry != null)
                {
                    //Utility.Log("Inside getItemRevisionQuery", logFilePath);

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
                        Utility.Log("getItemRevisionQuery: " + "executeQry is NULL: ", logFilePath);
                        return null;
                    }

                    Teamcenter.Services.Strong.Query._2007_06.SavedQuery.SavedQueryResults qryResult = executeQry.ArrayOfResults[0];
                    if (qryResult == null)
                    {
                        Utility.Log("getItemRevisionQuery: " + "qryResult is NULL: ", logFilePath);
                        return null;
                    }


                    // Assumption - Last Entry in this Query Result is the latest Revision... (Doubtful ??)
                    if (qryResult.Objects.Length > 0)
                    {

                        //setObjectPolicy();

                        String uid = qryResult.Objects[qryResult.Objects.Count() - 1].Uid;

                        if (dmService == null)
                        {
                            Utility.Log("getItemRevisionQuery: " + "dmService is NULL: ", logFilePath);
                            return null;
                        }

                        ServiceData sData = dmService.LoadObjects(new String[] { uid });

                        itemRevModelObject = (ModelObject)sData.GetPlainObject(0);

                        dmService.RefreshObjects(new ModelObject[] { itemRevModelObject });

                        //Utility.Log("Queryresult Count/Length ::" + qryResult.Objects.Count() + "/" + qryResult.Objects.Length, logFilePath);


                    }
                }

                else
                {
                    //Utility.Log("getItemRevisionQuery: " + "Item Revisions... query not found in Teamcenter.", logFilePath);
                }

                //Utility.Log( "Checking Item Revision in Teamcenter by Item ID & Rev ID process is completed.",logFilePath);


                return itemRevModelObject;
            }
            catch (Exception e)
            {
                Utility.Log("getItemRevisionQuery: " + "Exception in CheckItemIdBySysID functions: " + e.Message, logFilePath);
                Utility.Log("getItemRevisionQuery: " + "Exception in CheckItemIdBySysID functions: " + e.StackTrace, logFilePath);

            }
            return null;
        }
        public static void getNamedReferenceFile(Dataset ds, String StageDir, String logFilePath,
            Boolean parentItemRevDownloadFlag)
        {
            try
            {
                if (ds == null)
                {
                    Utility.Log("getNamedReferenceFile: " + "ds is" + " Empty", logFilePath);
                    return;
                }
                Utility.Log("getNamedReferenceFile: StageDir:" + StageDir, logFilePath);


                String fscHost = TCPropertyReader.get(Constants.TC_FSC_HOST);
                if (fscHost == null || fscHost.Equals("") == true)
                {
                    Utility.Log("getNamedReferenceFile: " + "fscHost is" + " Empty", logFilePath);
                    return;

                }
                String[] url = { fscHost };

                if (Teamcenter.ClientX.Session.getConnection() == null)
                {
                    Utility.Log("getNamedReferenceFile:" + "No Teamcenter Connection...", logFilePath);
                    return;
                }

                // String serverHost = TCPropertyReader.get(Constants.TC_SERVER_HOST);
                String serverHost = Constants.TC_SERVER_HOST;
                if (serverHost == null || serverHost.Equals("") == true)
                {
                    Utility.Log("getNamedReferenceFile: " + "serverHost is" + " Empty", logFilePath);
                    return;

                }

                FileManagementUtility fileUtility = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection(), serverHost, url, url, StageDir);
                DataManagementService dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());

                if (fileUtility == null)
                {
                    Utility.Log("getNamedReferenceFile: " + "fileUtility is" + " Empty", logFilePath);
                    return;
                }

                if (dmService == null)
                {
                    Utility.Log("getNamedReferenceFile: " + "dmService is" + " Empty", logFilePath);
                    return;
                }

                if (ds == null)
                {
                    Utility.Log("getNamedReferenceFile: " + "dataset is" + " Empty", logFilePath);
                    return;
                }

                //Teamcenter.Soa.Client.FileManagementUtility fileMangeUtil = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection(), "139.64.32.153", url, url, StageDir);
                ModelObject[] NRfiles = ds.Ref_list;
                Utility.Log("NRfiles count: " + NRfiles.Length, logFilePath);
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
                    Utility.Log("File is: " + DestinationFileFullPath, logFilePath);
                    if (File.Exists(DestinationFileFullPath) == true)
                    {
                        File.Delete(DestinationFileFullPath);
                    }
                    Object tempObj = new Object();
                    GetFileResponse res = fileUtility.GetFileToLocation(NRfile, DestinationFileFullPath, null, tempObj);

                    System.IO.FileInfo[] files = res.GetFiles();
                    Utility.Log("fileName: " + files[0].FullName, logFilePath);

                    for (int j = 0; j < files.Length; j++)
                    {
                        String path = files[j].FullName;
                        //Utility.Log("filePath: " + path, logFilePath);
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
                Utility.Log("Exception @ getNamedReferenceFile::" + e.Message, logFilePath);
                Utility.Log("Exception @ getNamedReferenceFile::" + e.StackTrace, logFilePath);
            }
        }




    }
}
