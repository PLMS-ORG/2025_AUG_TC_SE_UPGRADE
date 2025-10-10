using _3DAutomaticUpdate.controller;
using _3DAutomaticUpdate.model;
using _3DAutomaticUpdate.opInterfaces;
using _3DAutomaticUpdate.utils;
using SolidEdge.Framework.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate
{
    class ExcelSync
    {
        public static void ReadComponentTabFromExcel(Microsoft.Office.Interop.Excel.Application xlApp, String logFilePath)
        {
            try
            {
                Utility.Log("ReadMasterAssemblySheet: ", logFilePath);
                MasterAssemblyReader.ReadMasterAssemblySheet(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("ReadMasterAssemblySheet: " + ex.Message, logFilePath);
                return;
            }

            try
            {
                Utility.Log("RemoveSheet: ", logFilePath);
                ExcelRemoveComponent.RemoveSheet(xlApp, xlApp.ActiveWorkbook, logFilePath);
                xlApp.ActiveWorkbook.Save();
            }
            catch (Exception ex)
            {
                Utility.Log("RemoveSheet: " + ex.Message, logFilePath);
                return;
            }

            try
            {
                Utility.Log("RemoveComponentsFromFeatureTab: ", logFilePath);
                ExcelRemoveComponent.RemoveComponentsInFeatureTAB(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("RemoveComponentsFromFeatureTab: " + ex.Message, logFilePath);
                return;
            }


        }


        public static void SyncToSolidEdge(Microsoft.Office.Interop.Excel.Application xlApp, String topLineAssembly, String logFilePath)
        {

            Utility.Log("-----------------------------------------------------------------", logFilePath);
            Utility.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            try
            {
                //Utility.Log("Saving the Changes Done..", logFilePath, "INFO");
                xlApp.ActiveWorkbook.Save();

                Utility.Log("Reading Data from Excel..", logFilePath);
                ExcelData.readOccurenceVariablesFromTemplateExcelFast(xlApp, xlApp.ActiveWorkbook, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Exception: " + ex.Message, logFilePath);
                Utility.Log("Exception: " + ex.StackTrace, logFilePath);
                return;
            }

            //Utility.Log("Connecting to Solid Edge..", logFilePath, "INFO");
            //SE_SESSION.InitializeSolidEdgeSession(logFilePath);

            try
            {
                Utility.Log("Syncing Data to Solid Edge..", logFilePath);
                SolidEdgeFramework.Application application = null;
                SolidEdgeDocument document = null;
                application = SE_SESSION.getSolidEdgeSession();
                if (application == null)
                {
                    Utility.Log("Solid Edge Application is NULL", logFilePath);
                    return;
                }
                document = (SolidEdgeDocument)(SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;

                if (document == null)
                {
                    Utility.Log("Solid Edge Document is NULL", logFilePath);
                    return;
                }
                topLineAssembly = document.FullName;
                Utility.Log("topLineAssembly: " + topLineAssembly, logFilePath);
                SolidEdgeInterface.SolidEdgeSync(topLineAssembly, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Exception: " + ex.Message, logFilePath);
                Utility.Log("Exception: " + ex.StackTrace, logFilePath);
                return;
            }



            Utility.Log("-----------------------------------------------------------------", logFilePath);
            Utility.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            Utility.Log("Completed Syncing Data To Solid Edge..", logFilePath);
            //MessageBox.Show("Sync to Solid Edge Completed");
            //xlApp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            //xlApp = null;
        }



        public static void SyncFeaturesToSolidEdge(Microsoft.Office.Interop.Excel.Application xlApp, String topLineAssembly, String logFilePath)
        {

            String xlFilePath = xlApp.ActiveWorkbook.FullName;
            Utility.Log("SolidEdgeSetFeatureState: xlFilePath: " + xlFilePath, logFilePath);
            Utility.Log("SolidEdgeSetFeatureState: readFeaturesFromTemplateExcel: " + xlFilePath, logFilePath);
            ExcelReadFeatures.readFeaturesFromTemplateExcel(xlApp, xlFilePath, logFilePath);
            Utility.Log("SolidEdgeSetFeatureState: getFeatureLinesList: " + xlFilePath, logFilePath);
            List<FeatureLine> updatedFsList = ExcelReadFeatures.getFeatureLinesList();
            try
            {
                Utility.Log("SolidEdgeSetFeatureState: setFeatures: " + System.DateTime.Now.ToString(), logFilePath);
                SolidEdgeSetFeatureState SetFeature = new SolidEdgeSetFeatureState();
                SetFeature.SolidEdgeFeatureSyncFromExcel(topLineAssembly, updatedFsList, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);
                return;
            }
            Utility.Log("SyncFeaturesToSolidEdge: Completed: " + System.DateTime.Now.ToString(), logFilePath);
        }

    }
}
