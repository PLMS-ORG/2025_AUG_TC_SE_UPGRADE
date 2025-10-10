using _3DAutomaticUpdate.model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate.controller
{
    class ExcelReadFeatures
    {
        public static Dictionary<String, List<FeatureLine>> featureDictionary = new Dictionary<string, List<FeatureLine>>();
        public static List<FeatureLine> fsList = new List<FeatureLine>();

        public static Dictionary<String, List<FeatureLine>> getFeatureDictionary()
        {
            featureDictionary = Utility.BuildFeatureDictionary(fsList, "");
            return featureDictionary;
        }

        public static List<FeatureLine> getFeatureLinesList()
        {
            return fsList;
        }

        public static void readFeaturesFromTemplateExcel(Microsoft.Office.Interop.Excel._Application xlApp, String xlFilePath, String logFilePath)
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


            FileInfo f = new FileInfo(xlFilePath);

            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                //Utility.Log("File Already Exists", logFilePath);
                xlApp.DisplayAlerts = false;
                xlWorkbook = xlApp.ActiveWorkbook;
            }
            else
            {
                Utility.Log("file does not Exist: " + xlFilePath, logFilePath);
                return;
            }
            if (xlWorkbook == null)
            {
                Utility.Log("xlWorkBook is NULL", logFilePath);
                return;
            }

            //Utility.Log("WorkSheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {

                List<Variable> variableListForSheet = new List<Variable>();

                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //Utility.Log(sheet.Name, logFilePath);
                    readFeatures(sheet, logFilePath);
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

            //xlApp.Visible = true;
            //xlApp.UserControl = false; 

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            if (xlApp != null) xlApp.DisplayAlerts = false;
            //Utility.Log("----------------------------------------------------------", logFilePath);
            //quit and release
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);
            //xlApp = null;
            return;

        }

        private static void readFeatures(Microsoft.Office.Interop.Excel.Worksheet sheet, String logFilePath)
        {
            Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
            //Utility.Log(xlRange.Rows.Count.ToString(), logFilePath);
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;
                try
                {
                    //Utility.Log("Iteration: " + i.ToString(), logFilePath);

                    FeatureLine f = new FeatureLine();
                    if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                    {
                        try
                        {

                            f.PartName = xlRange.Cells[i, 6].Value2;
                            //Utility.Log("PartName: " + f.PartName, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utility.Log("PartName" + ex.Message, logFilePath);
                        }
                    }
                    if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                    {
                        try
                        {

                            f.EdgeBarName = xlRange.Cells[i, 7].Value2;
                            //Utility.Log("EdgeBarName: " + f.PartName, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utility.Log("EdgeBarName" + ex.Message, logFilePath);
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
                            Utility.Log("FeatureName" + ex.Message, logFilePath);
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
                            Utility.Log("SystemName" + ex.Message, logFilePath);

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
                            Utility.Log("Formula" + ex.Message, logFilePath);
                        }
                    }

                    //if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                    if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                    {
                        try
                        {
                            //f.IsFeatureEnabled = xlRange.Cells[i, 5].Value2;
                            f.SuppressionEnabled = xlRange.Cells[i, 8].Value2;
                            //Utility.Log("ExcelData: " + xlRange.Cells[i, 8].Value2, logFilePath);
                            //Utility.Log("SuppressionEnabled: " + f.SuppressionEnabled, logFilePath);
                            if (f.SuppressionEnabled != null && f.SuppressionEnabled.Equals("") == false)
                            {
                                if (f.SuppressionEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    f.IsFeatureEnabled = "N";
                                }
                                else
                                {
                                    f.IsFeatureEnabled = "Y";
                                }
                            }
                            //Utility.Log("IsFeatureEnabled: " + f.IsFeatureEnabled, logFilePath);
                            // 02 - OCT, No need to Set SuppressionEnabled Property. If IsFeatureEnabled Property is Set, SuppressionEnabled Property is Automatically Retrieved Out.
                            // 28 - OCT, When Users change SuppressionEnabled, FeatureEnabled to be Set Automatically
                        }
                        catch (Exception ex)
                        {
                            Utility.Log("SuppressionEnabled" + ex.Message, logFilePath);


                        }
                    }

                    fsList.Add(f);
                }
                catch (Exception ex)
                {
                    Utility.Log(ex.Message, logFilePath);
                }
            }

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;
        }
    }




}
