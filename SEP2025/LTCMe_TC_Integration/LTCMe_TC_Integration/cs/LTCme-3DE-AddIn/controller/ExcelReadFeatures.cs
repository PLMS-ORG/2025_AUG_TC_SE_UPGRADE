/**************************************************************************
 *  3DE International FZE, 2018
 *  Murali - 02-OCT-18 - Changes for SuppressionEnabled Property
 *  ***********************************************************************
 */

using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DemoAddInTC.controller
{
    class ExcelReadFeatures
    {
        public static Dictionary<String, List<FeatureLine>> featureDictionary = new Dictionary<string, List<FeatureLine>>();
        public static List<FeatureLine> fsList = new List<FeatureLine>();

        public static Dictionary<String, List<FeatureLine>> getFeatureDictionary()
        {
            featureDictionary = Utlity.BuildFeatureDictionary(fsList, "");
            return featureDictionary;
        }

        public static List<FeatureLine> getFeatureLinesList()
        {
            return fsList;
        }

        public static void readFeaturesFromTemplateExcel(String xlFilePath, String logFilePath)
        {
            // If Not Cleared, Features get Added in the View.
            if (featureDictionary != null)
            {
                featureDictionary.Clear();
            }
            if (fsList != null)
            {
                fsList.Clear();
            }
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            if (f.Exists == true)
            {
                workbooks = xlApp.Workbooks;
                Utlity.Log("File Already Exists", logFilePath);
                xlApp.DisplayAlerts = false;
                //xlWorkbook = workbooks.Open(xlFilePath);
                try
                {
                    xlWorkbook = workbooks.Open(xlFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = workbooks.Open(xlFilePath, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Utlity.Log("file does not Exist: " + xlFilePath, logFilePath);
                return;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
           
            //Utlity.Log("WorkSheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {

                List<Variable> variableListForSheet = new List<Variable>();
                
                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //Utlity.Log(sheet.Name, logFilePath);
                    //readFeatures(sheet,logFilePath);
                    readFeaturesFast(sheet, logFilePath);

                }
                else
                {
                     Marshal.ReleaseComObject(sheet);
                    continue;
                }

                Marshal.ReleaseComObject(sheet);
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Close(true);


            Marshal.ReleaseComObject(sheets);
            sheets = null;

            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            Marshal.ReleaseComObject(workbooks);
            workbooks = null;

            Utlity.Log("----------------------------------------------------------", logFilePath);
            xlApp.DisplayAlerts = true;
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            return;

        }

        private static void readFeatures(Microsoft.Office.Interop.Excel.Worksheet sheet, String logFilePath)
        {
            Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
             //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    if (i == 1)
                        continue;
                    try
                    {
                        //Utlity.Log("Iteration: " + i.ToString(), logFilePath);

                        FeatureLine f = new FeatureLine();
                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {
                            try
                            {

                                f.PartName = xlRange.Cells[i, 6].Value2;
                                //Utlity.Log("PartName" + f.PartName, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("PartName" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        {
                            try
                            {

                                f.EdgeBarName = xlRange.Cells[i, 7].Value2;
                                //Utlity.Log("EdgeBarName" + f.PartName, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("EdgeBarName" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2] != null)
                        {
                            try
                            {
                                f.FeatureName = xlRange.Cells[i, 2].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("FeatureName" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        {
                            try
                            {
                                f.SystemName = xlRange.Cells[i, 3].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("SystemName" + ex.Message, logFilePath);
                                
                            }
                        }
                        if (xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4] != null)
                        {
                            try
                            {
                                f.Formula = xlRange.Cells[i, 4].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Formula" + ex.Message, logFilePath);
                            }
                        }

                        //if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                        if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                        {
                            try
                            {
                                //f.IsFeatureEnabled = xlRange.Cells[i, 5].Value2;
                                String SuppressionEnabled = xlRange.Cells[i, 8].Value2;
                                Utlity.Log("ExcelData: " + xlRange.Cells[i, 8].Value2, logFilePath);
                                Utlity.Log("SuppressionEnabled: " + SuppressionEnabled, logFilePath);
                                if (SuppressionEnabled != null && SuppressionEnabled.Equals("") == false)
                                {
                                    if (SuppressionEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        f.IsFeatureEnabled = "N";
                                    }
                                    else
                                    {
                                        f.IsFeatureEnabled = "Y";
                                    }
                                }
                                Utlity.Log("IsFeatureEnabled: " + f.IsFeatureEnabled, logFilePath);
                                // 02 - OCT, No need to Set SuppressionEnabled Property. If IsFeatureEnabled Property is Set, SuppressionEnabled Property is Automatically Retrieved Out.
                                // 28 - OCT, When Users change SuppressionEnabled, FeatureEnabled to be Set Automatically
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("SuppressionEnabled" + ex.Message, logFilePath);


                            }
                        }
                        
                        fsList.Add(f);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                    }
        }

                Marshal.ReleaseComObject(xlRange);
                xlRange = null;
    }


        private static void readFeaturesFast(Microsoft.Office.Interop.Excel.Worksheet sheet, String logFilePath)
        {
            Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
            //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);

            object[,] values = (object[,])xlRange.Value2;

            for (int i = 1; i <= values.GetLength(0); i++)
            //for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;
                try
                {
                    //Utlity.Log("Iteration: " + i.ToString(), logFilePath);

                    FeatureLine f = new FeatureLine();
                    //if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                    if (values[i,6] !=null)
                    {
                        try
                        {
                            f.PartName = Convert.ToString(values[i, 6]);
                            //f.PartName = xlRange.Cells[i, 6].Value2;
                            //Utlity.Log("PartName" + f.PartName, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("PartName" + ex.Message, logFilePath);
                        }
                    }
                    //if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                    if (values[i, 7] != null)
                    {
                        try
                        {

                            f.EdgeBarName = Convert.ToString(values[i, 7]);
                            //f.EdgeBarName = xlRange.Cells[i, 7].Value2;
                            //Utlity.Log("EdgeBarName" + f.PartName, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("EdgeBarName" + ex.Message, logFilePath);
                        }
                    }
                   // if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2] != null)
                    if (values[i, 2] != null)
                    {
                        try
                        {
                            f.FeatureName = Convert.ToString(values[i, 2]);
                            //f.FeatureName = xlRange.Cells[i, 2].Value2;
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("FeatureName" + ex.Message, logFilePath);
                        }
                    }
                    //if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                    if (values[i, 3] != null)
                    {
                        try
                        {
                            //f.SystemName = xlRange.Cells[i, 3].Value2;
                            f.SystemName = Convert.ToString(values[i, 3]);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("SystemName" + ex.Message, logFilePath);

                        }
                    }
                    //if (xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4] != null)
                    if (values[i, 4] != null)
                    {
                        try
                        {
                            //f.Formula = xlRange.Cells[i, 4].Value2;
                            f.Formula = Convert.ToString(values[i, 4]);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("Formula" + ex.Message, logFilePath);
                        }
                    }

                    //if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                    //if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                    if (values[i, 8] != null) // Suppression Enabled
                    {
                        try
                        {
                            String SuppressionEnabled = Convert.ToString(values[i, 8]);
                            //f.IsFeatureEnabled = xlRange.Cells[i, 5].Value2;
                            //String SuppressionEnabled = xlRange.Cells[i, 8].Value2;
                            //Utlity.Log("ExcelData: " + xlRange.Cells[i, 8].Value2, logFilePath);
                            //Utlity.Log("SuppressionEnabled: " + SuppressionEnabled, logFilePath);
                            if (SuppressionEnabled != null && SuppressionEnabled.Equals("") == false)
                            {
                                if (SuppressionEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    f.IsFeatureEnabled = "N";
                                }
                                else
                                {
                                    f.IsFeatureEnabled = "Y";
                                }
                            }
                            //Utlity.Log("IsFeatureEnabled: " + f.IsFeatureEnabled, logFilePath);
                            // 02 - OCT, No need to Set SuppressionEnabled Property. If IsFeatureEnabled Property is Set, SuppressionEnabled Property is Automatically Retrieved Out.
                            // 28 - OCT, When Users change SuppressionEnabled, FeatureEnabled to be Set Automatically
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("SuppressionEnabled" + ex.Message, logFilePath);


                        }
                    }

                    fsList.Add(f);
                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }
            }

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;
        }
}



    }


        
