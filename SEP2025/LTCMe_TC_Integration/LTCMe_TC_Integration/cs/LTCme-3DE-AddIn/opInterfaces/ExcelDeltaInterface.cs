using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

// 28 SEPT - Added to Just change the Values in the XL sheet & nothing Else.
namespace DemoAddInTC.opInterfaces
{
    class ExcelDeltaInterface
    {
        public static bool SaveDeltaToXL(String outputXLfileName,
           String logFilePath)
        {
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;            

            Utlity.Log("Inside SaveDeltaToXL", logFilePath);
            Dictionary<String, List<Variable>> variableDictionary = SolidEdgeData1.getVariablesDictionaryDetails();
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utlity.Log("variableDictionary is NULL,", logFilePath);
                return false;
            }
            xlApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            FileInfo f = new FileInfo(outputXLfileName);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                Utlity.Log("File Already Exists,", logFilePath);
                //xlWorkbook = xlApp.Workbooks.Open(outputXLfileName);
                try
                {
                    xlWorkbook = xlApp.Workbooks.Open(outputXLfileName);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = xlApp.Workbooks.Open(outputXLfileName, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;
                return false;
            }
            if (xlWorkbook == null)
            {
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return false;
            }
            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in xlWorkbook.Worksheets)
            {
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        Utlity.Log("Skipping Custom LTC Sheet: " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }

                Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                {
                    // 10 - SEPT - NOT READING SHEETS WHICH ARE HIDDEN. PERFORMANCE OVER HEAD.
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                else
                {
                    try
                    {
                        List<Variable> variablesList = null;
                        variableDictionary.TryGetValue(sheet.Name, out variablesList);
                        if (variablesList != null && variablesList.Count != 0)
                        {
                            Utlity.Log(sheet.Name + ":::" + variablesList.Count.ToString(), logFilePath);
                            WriteSheet(xlWorkbook, sheet, sheet.Name, variablesList, logFilePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        //Marshal.ReleaseComObject(xlWorkbook);
                        //xlWorkbook = null;
                        //Marshal.ReleaseComObject(workbooks);
                        //workbooks = null;
                        Marshal.ReleaseComObject(sheet);
                        Utlity.Log("SaveDeltaToXL: Exception " + ex.Message, logFilePath);
                        continue;
                    }
                }


                Marshal.ReleaseComObject(sheet);
            }

            Marshal.ReleaseComObject(sheets);
            sheets = null;


            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;


            Marshal.ReleaseComObject(workbooks);
            workbooks = null;

            //quit and release
            if (xlApp != null)  xlApp.DisplayAlerts = true;
            if (xlApp != null) xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

            return true;
        }

        // 12- SEPT Writes the Modified Value from variableArr List to the Excel Sheet.
        // SyncTE would update the Value and Reevaluate the Expressions in SE and Write back the Data into XL.
        private static void WriteSheet(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel.Worksheet sheet, String ocurrenceName, List<Variable> variableArr, String logFilePath)
        {


            if (sheet == null)
            {
                Utlity.Log("WriteSheet: sheet is Empty" + ocurrenceName, logFilePath);
                return;
            }

            sheet.Name = ocurrenceName;
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;


            //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;

                // NAME
                if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                {
                    try
                    {
                        String name = xlRange.Cells[i, 1].Value2;
                        if (name != null && name.Equals("") == false)
                        {
                            var element = variableArr.Find(var => var.name.Equals(name, StringComparison.OrdinalIgnoreCase));
                            if (element != null)
                            {
                                // 7 October 2018, 
                                String Formula = xlRange.Cells[i, 3].Formula;
                               // Utlity.Log("Formula: " + Formula, logFilePath);
                                if (Formula != null && Formula.Equals("") == false)
                                {
                                    if (Formula.StartsWith("=") == true)
                                    {
                                        Utlity.Log("Formula: " + Formula, logFilePath);
                                    }
                                    else
                                    {
                                        // SET VALUE
                                        xlRange.Cells[i, 3].Value2 = element.value;
                                    }
                                }
                            }
                            else
                            {
                                Utlity.Log("WriteSheet: " + "Could Not Find " + name, logFilePath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("name" + ex.Message, logFilePath);
                    }
                }

            }
            xlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

            //Marshal.ReleaseComObject(sheet);
            //sheet = null;


        }


        // 27 - OCT, Feature Sync from Excel would Fail if Variable Parts are not Renamed in Features TAB.
        public static void RenameVariablePartsInFeaturesTab(String assemblyFileName, String FolderToPublish, List<String> variablePartsList, String Suffix, String logFilePath)
        {
            String FileName = System.IO.Path.GetFileName(assemblyFileName);
            if (FileName == null || FileName.Equals("") == true)
            {
                Utlity.Log("RenameSheetNamesForVariableParts: FileName is Empty..", logFilePath);
                return;
            }
            String xlFile = System.IO.Path.ChangeExtension(FileName, ".xlsx");
            if (xlFile == null || xlFile.Equals("") == true)
            {
                Utlity.Log("RenameSheetNamesForVariableParts: xlFile is Empty..", logFilePath);
                return;
            }
            String xlFilePath = System.IO.Path.Combine(FolderToPublish, xlFile);

            if (variablePartsList == null || variablePartsList.Count == 0)
            {
                Utlity.Log("RenameSheetNamesForVariableParts: NO VARIABLE PARTS FOUND..", logFilePath);
                return;
            }

            if (Suffix == null || Suffix.Equals("") == true)
            {
                Utlity.Log("RenameSheetNamesForVariableParts: Suffix is Empty..", logFilePath);
                return;
            }

            if (xlFilePath == null || xlFilePath.Equals("") == true)
            {
                Utlity.Log("RenameSheetNamesForVariableParts: XLFilePath is Empty..", logFilePath);
                return;
            }

            UpdateOccurrenceNameInFeaturesTAB(xlFilePath, variablePartsList, Suffix, logFilePath);
            return ;
        }


        public static void UpdateOccurrenceNameInFeaturesTAB(String xlFilePath, List<String> variablePartsList, String Suffix,  String logFilePath)
        {
            Utlity.Log("UpdateOccurrenceNameInFeaturesTAB: XLFilePath: " + xlFilePath, logFilePath);
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            //xlApp.WindowState = XlWindowState.xlNormal;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                //xlWorkbook = xlApp.Workbooks.Open(xlFilePath);
                try
                {
                    xlWorkbook = xlApp.Workbooks.Open(xlFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = xlApp.Workbooks.Open(xlFilePath, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Utlity.Log("File Does Not Exist..,", logFilePath);
                return;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
           
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
                    xlWorkSheet.Activate();
                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i == 1)
                            continue;

                        if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                        {
                            try
                            {

                                String PartName = (String)xlRange.Cells[i, 1].Value2;
                                //Utlity.Log("PartName: " + PartName, logFilePath);
                                if (PartName != null && PartName.Equals("") == false)
                                {
                                    if (variablePartsList.Contains(PartName) == true)
                                    {
                                        String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(PartName);
                                        String extn = System.IO.Path.GetExtension(PartName);
                                        String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                                        try
                                        {
                                            xlRange.Cells[i, 1].Value2 = newPartFileName;
                                        }
                                        catch (Exception ex)
                                        {
                                            Utlity.Log("PartName: " + ex.Message, logFilePath);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("PartName: " + ex.Message, logFilePath);
                            }

                        }

                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {
                            try
                            {

                                String PartName = (String)xlRange.Cells[i, 6].Value2;
                                //Utlity.Log("PartName: " + PartName, logFilePath);
                                if (PartName != null && PartName.Equals("") == false)
                                {
                                    if (variablePartsList.Contains(PartName) == true)
                                    {
                                        String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(PartName);
                                        String extn = System.IO.Path.GetExtension(PartName);
                                        String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                                        try
                                        {
                                            xlRange.Cells[i, 6].Value2 = newPartFileName;
                                        }
                                        catch (Exception ex)
                                        {
                                            Utlity.Log("PartName: " + ex.Message, logFilePath);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("PartName: " + ex.Message, logFilePath);
                            }
                        }
                    }
                    Marshal.ReleaseComObject(sheet);
                }
            }



            xlApp.Visible = false;
            xlApp.UserControl = false;
            if (f.Exists == true)
            {
                Utlity.Log(xlFilePath + " Being Saved...", logFilePath);
                xlWorkbook.Save();
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            if (workbooks != null)
            {
                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;
            }
            if (xlApp != null) xlApp.DisplayAlerts = true;
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            Utlity.Log("Completed Saving Template Excel", logFilePath);
           
        }


       
    }
}
