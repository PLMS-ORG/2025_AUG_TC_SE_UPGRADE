using Creo_TC_Live_Integration.TcDataManagement;
using DemoAddInTC.controller;
using DemoAddInTC.utils;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using Teamcenter.Net.TcServerProxy.Admin;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;

namespace DemoAddInTC.se
{
    class CheckInOut
    {
 
        public static void RunCheckInCheckOutMethod(String option, SolidEdgeFramework.Application application)
        {
            // SolidEdgeFramework.Application application = null;
            SolidEdgeDocument document = null;

            //Connect to running Solid Edge Instance
            //application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            if (application == null)
            {
                MessageBox.Show("Solid Edge Application is NULL");
                return;
            }

            SE_SESSION.setSolidEdgeSession(application);
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

            if (document == null)
            {
                MessageBox.Show("Solid Edge document is NULL");
                return;
            }
            String assemblyFileName = document.FullName;
            SolidEdgeData1.setAssemblyFileName(assemblyFileName);
            //String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
            String stageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + "_" + option + ".txt");

            Utlity.Log("-----------------------------------------------------------------", logFilePath, option);
            Utlity.Log("Run " + option + " Utility Started @ " + System.DateTime.Now.ToString(), logFilePath, option);

            Utlity.Log("SEEC Login..", logFilePath, option);
            SEECAdaptor.LoginToTeamcenter(logFilePath);
            string bStrCurrentUser = null;
            SEECAdaptor.getSEECObject().GetCurrentUserName(out bStrCurrentUser);

            String password = bStrCurrentUser;

            Utlity.Log("Logging into TC..for Sanitize XL (PostClone)", logFilePath);
            Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath);
            Utlity.Log("ID=" + bStrCurrentUser, logFilePath);
            Utlity.Log("Group=Engineering", logFilePath);
            Utlity.Log("Role=Designer", logFilePath);
            TcAdaptor.login(bStrCurrentUser, password, "Engineering", "Designer", logFilePath);
            Utlity.Log("Initializing TC Services..", logFilePath);
            TcAdaptor.TcAdaptor_Init(logFilePath);

            //Utlity.Log("Logging into TC..for " + option, logFilePath, option);
            //TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
            //Utlity.Log("Initializing TC Services..", logFilePath, option);
            //TcAdaptor.TcAdaptor_Init(logFilePath);
            //SEECAdaptor.LoginToTeamcenter(logFilePath);

            String RevID = SEECAdaptor.getRevisionID(assemblyFileName);
            String itemID = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
            String cachePath = SEECAdaptor.GetPDMCachePath();

            if (cachePath == null || cachePath == "" || cachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                cachePath = Path.GetDirectoryName(assemblyFileName);
            }

            String XlTemplateFile = Path.ChangeExtension(assemblyFileName, ".xlsx");

            Utility.Log(option + " Excel: getItemRevisionQuery..", logFilePath);
            ItemRevision itemRevMO = (ItemRevision)DownloadDatasetNamedReference.getItemRevisionQuery(itemID, RevID, logFilePath);

            if (itemRevMO == null)
            {
                Utility.Log(option + " Excel: item REV model Object is NULL/Empty", logFilePath);
                return;

            }

            Utility.Log(option + " Excel: isDataSetAvailable..", logFilePath);
            ModelObject dataSetMo = TcAdaptor.isDataSetAvailable(itemRevMO, "IMAN_specification", "MS ExcelX");
       
            if (dataSetMo == null)
            {
                Utility.Log(option + " Excel:..Excel Dataset is not available under the Item Revision..", logFilePath);
                return;
            }
            if (option.Equals("CheckOut", StringComparison.OrdinalIgnoreCase) == true)
            {
                 ModelObject dsMo = TcAdaptor.checkOutModelObject(dataSetMo, logFilePath);
                    if (dsMo != null)
                    {
                        Utility.Log("Excel Dataset is Checked Out...", logFilePath);
                        return;
                    }
                }           

            if (option.Equals("CheckIn", StringComparison.OrdinalIgnoreCase) == true)
            {
                Utility.Log("XlTemplateFile: " + XlTemplateFile, logFilePath);
                bool openFlag = IsExcelFileOpen(XlTemplateFile);
                if (openFlag == false) 
                {

                    TcAdaptor.uploadExcelToTC(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, XlTemplateFile, logFilePath);
  
                }
                else
                {
                    Utility.Log("XlTemplateFile is open, Close it before Check In.." + XlTemplateFile, logFilePath);
                    return;
                }

                bool flag = TcAdaptor.checkInModelObject(dataSetMo);
                if (flag == false)
                {
                    Utility.Log("Excel Dataset is NOT Checked In...", logFilePath);
                    return;
                }
            }

            Utility.Log("TC logout completed", logFilePath);
            TcAdaptor.logout(logFilePath);
        }

      public static bool IsExcelFileOpen(string path)
        {
            FileStream stream = null;
            try
            {
                stream = File.Open(path, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            }
            catch (IOException ex)
            {
                if (ex.Message.Contains("being used by another process"))
                    return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            return false;
        }

        //=========================================================================================
        public static bool isExcelDSCheckedOut(ModelObject dataSetMo, String logFilePath)
        {
            String dsUID = dataSetMo.Uid;

            Utility.Log("INFO : dsUID " + dsUID, logFilePath);
            String[] arr = new String[] { dsUID };

            ServiceData sData = TcAdaptor.dmService.LoadObjects(arr);

            if (sData == null)
            {
                Utility.Log("Error : ServiceData is null", logFilePath);
                return false;
            }

            Dataset dsMObj = (Dataset)sData.GetPlainObject(0);

            if (dsMObj == null)
            {
                Utility.Log("Error : Excel Dataset is null", logFilePath);
                return false;
            }

            Utility.Log(" Excel Dataset is Checked Out: " + dsMObj.Checked_out, logFilePath);

            if (dsMObj.Checked_out.CompareTo("Y") == 0)
            {
                Utility.Log("INFO : Excel Dataset is Checked Out: ", logFilePath);
                return true;
            }
            else
            {
                Utility.Log("INFO : Excel Dataset is NOT alraedy  Checked Out: ", logFilePath);
                return false;
            }
        }
        //===============================================================================
    }


}


