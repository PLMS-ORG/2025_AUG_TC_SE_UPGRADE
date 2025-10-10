using System;
using System.Runtime.InteropServices;

using System.IO;
using CreoToTc.Utils;
using Log;
using AddToTc.CDAT_BulkUploader;
using Creo_TC_Live_Integration.TeamCenter;
using DemoAddInTC.utils;
using Creo_TC_Live_Integration.TcDataManagement;
using CTD;
using DemoAddInTC.se;
using DemoAddInTC.controller;
using Teamcenter.Soa.Client.Model.Strong;
using Teamcenter.Soa.Client.Model;
using AddToTc.CTD;
using System.Collections.Generic;

namespace CDAT_BulkUploader
{
    class Program
    {
        //static void Main (string[] args)
        //{
        //    try
        //    {
        //        string userName = "dcproxy";
        //        string password = "dcproxy";
        //        string group = "";
        //        string role = "";
        //        string URL = "http://aewsrv-tcsr02/tc/"; //prod

        //        Constants.TC_SERVER_HOST = "http://aewsrv-tcsr02/tc"; //-server            
        //        Constants.TC_FSC_HOST = "http://aewsrv-tcsr02:4544";

        //        if (TCUtils.tc_LogIn(userName, password, group) == false)
        //        {
        //            log.write(logType.ERROR, "Login failed");
        //            Constants.globalError.Add("Teamcenter login failed.");
        //            return;
        //        }

        //        Tc_Services tcService = new Tc_Services();

        //        // Set object policiese
        //        if (ObjectPolicy.setObjectPolicy_General())
        //            log.write(logType.INFO, "setObjectPolicy_General successful");
        //        else
        //            log.write(logType.ERROR, "setObjectPolicy_General failed..");

        //        Test_Function();

        //        if (Tc_Services.ss != null)
        //        {
        //            Tc_Services.ss.Logout();

        //        }
        //        else
        //        {
        //            log.write(logType.WARNING, "Not able to logout.");
        //        }
        //    } catch (Exception ex)
        //    {
        //        Constants.globalError.Add("Uploading failed.");

        //        log.writeException(ex, "main");
        //    }

        //}


        static void Main(string[] args)
        {
            try
            {
                if (args == null || args.Length != 7)
                {
                    Console.WriteLine("PLease provide valid arguments");
                    Utlity.Log("PLease provide valid arguments", string.Empty);
                    return;
                }
                if (args[2] == null)
                {
                    Utlity.Log("Please provide valid stage Dir", string.Empty);
                    return;
                }

                log.logFile = Path.Combine(args[2] + "log.txt");

                if (args[0] == null || args[1] == null)
                {
                    log.write(logType.ERROR, "Please provide valid ItemRevID");
                    return;
                }

                    string stageDir = args[2];
                    String bstrItemID = args[0];
                    String bstrItemRevID = args[1];
                    if (bstrItemID.Trim().Equals("") == true)
                        throw new Exception("Item ID from selected node is blank");
                    if (bstrItemRevID.Trim().Equals("") == true)
                        throw new Exception("Revision ID from selected node is blank");

                    log.write(logType.INFO, "ItemID = " + bstrItemID);
                    log.write(logType.INFO, "ItemRevID = " + bstrItemRevID);
                    log.write(logType.INFO, "");

                if (args[3] == null || args[4] == null)
                {
                    Utlity.Log("Please provide TC userName and Password in 3th and 4th args", string.Empty);
                    return;
                }

                if (args[5] == null || args[6] == null)
                {
                    Utlity.Log("Please provide TC URL and TC FSC HOST in 5th  and 6th args", string.Empty);
                    return;
                }

                string userName = args[3]; //dcproxy
                    string password = args[4];  //dcproxy
                    string group = "";
                    string role = "";
                    string URL = args[5];
                           
                    string newItemID = ""; // pallet asm 
                    string newItemRevID = ""; // pallet asm
                    string cachePath = ""; // where do you want the XL to be downloaded ?
                    Dictionary<String, string> ItemAndRevIdDictionary = null;
                    (newItemID, newItemRevID, cachePath, ItemAndRevIdDictionary) = SEECCTD.perform_ctd(bstrItemID, bstrItemRevID,
                                                                               userName, password, group, role, URL);

                write_derivative_info_to_text_file(newItemID, newItemRevID, stageDir);
                Constants.TC_SERVER_HOST = args[5]; //test server            
                Constants.TC_FSC_HOST = args[6];
      

                if (TCUtils.tc_LogIn(userName, password, group) == false)
                {
                    log.write(logType.ERROR, "Login failed");
                    Constants.globalError.Add("Teamcenter login failed.");
                    return;
                }

                Tc_Services tcService = new Tc_Services();

                // Set object policiese
                if (ObjectPolicy.setObjectPolicy_General())
                    log.write(logType.INFO, "setObjectPolicy_General successful");
                else
                    log.write(logType.ERROR, "setObjectPolicy_General failed..");

                // Call your tc functions
                // AVM - 16/03/2024
                DownloadDatasetNamedReference.RetrieveItemRevMOAndDownloadDatasetNR(newItemID, newItemRevID,
                cachePath, true, true);
                //-------------------- SOA Calls to TC

                // AVM - 16/03/2024
                String assemblyFileName = Path.Combine(cachePath, newItemID + ".asm");

                if (File.Exists(assemblyFileName) == false)
                {
                    log.write(logType.WARNING, "assemblyFileName does not exist: " + assemblyFileName);
                    return;
                }

                // read the assembly information by traversing till the child.
                // sanitize the input XL.
                // Upload the Excel back to Derivative Item Revision in Teamcenter.
                SanitizeXL_PostClone.read_all_items_in_cache(assemblyFileName, userName, password, group,stageDir);
                // AVM - 16/03/2024
                //-------------------- SOA Calls to TC
                TCUtils.PostCloneCleanUpExcelDataSet(newItemID,newItemRevID);

               if (Tc_Services.ss!=null)
               {
                   Tc_Services.ss.Logout();
                   
               }
               else
               {
                   log.write(logType.WARNING, "Not able to logout." );
               }
               }
               catch (Exception ex)
               {
                   Constants.globalError.Add("Uploading failed.");

                   log.writeException(ex, "main");
               }
            //-------------------- SOA Calls to TCs

        }

        // This text file will be used to send Email to TC User.
        private static void write_derivative_info_to_text_file(string newItemID, string newItemRevID,String stageDir)
        {
           
            if (newItemID == "" && newItemRevID == "")
            {
                log.write(logType.ERROR, "Create template derivative failed..");
                return;
            }
            else
                log.write(logType.INFO, "Derivative is created..");

            //write deruvativeItemId & derivativeRevID in output.txt
            string outputTxtPath =  Path.Combine(stageDir, "Output.txt");
          
            if (System.IO.File.Exists(outputTxtPath))
                System.IO.File.Delete(outputTxtPath);
            
            StreamWriter sw = new StreamWriter(outputTxtPath, false);
            sw.WriteLine(newItemID + "~" + newItemRevID);
            sw.Close();
        }


        //======================================================================================================
        //private static void Test_Function()
        //{
        //    String cachePath = "C:\\ProgramData\\";
            
        //    String xlFile = Path.Combine(cachePath, "01858037.xlsx");
        //    if (File.Exists(xlFile) == true)
        //    {
        //        log.write(logType.INFO, "Test_Function: Upload the excel back to Teamcenter: " + xlFile);
        //        TCUtils.uploadExcelToTC_Test("01858037", "A", xlFile);

        //    }
        //    else
        //    {
        //        log.write(logType.INFO, "Test_Function: " + "no xl file is found in cache");
        //    }
        //}
    }

    
}
