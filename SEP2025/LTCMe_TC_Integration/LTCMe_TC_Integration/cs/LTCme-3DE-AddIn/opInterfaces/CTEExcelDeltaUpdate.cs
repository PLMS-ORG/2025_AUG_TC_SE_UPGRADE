using DemoAddInTC.model;
using DemoAddInTC.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

// 18 - November - Added By Murali - 
namespace DemoAddInTC.opInterfaces
{
    class CTEExcelDeltaUpdate
    {
        //Step - 1 Open the XL and Read the Part Sheets.
        //Step- 2, Add/Delete/Hide Sheets - If Part is Added/Removed from Template.
        //Step-3, Open the Sheet, Read the Variable Rows, Add/Delete/Update Rows -- Based on Updated Data in Template
        //Step-4, Open the "Component" Sheet & Add/Update/Delete -- Based on Updated Component Data in Template
        //Step-5, Open the "Features" sheet & Add/Update/Delete -- Based on Updated Features Data in Template
        public static bool SaveDeltaToXL(String outputXLfileName, List<Variable> variablesList,
            Dictionary<String, bool> partEnablementDictionary, String logFilePath,
            String Option, List<FeatureLine> featureLineList)
        {
            Utlity.Log("Starting SaveDeltaToXL: " + outputXLfileName, logFilePath);
            if (partEnablementDictionary == null)
            {
                Utlity.Log("partEnablementDictionary is NULL,", logFilePath);
                return false;
            }
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = XlWindowState.xlNormal;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            FileInfo f = new FileInfo(outputXLfileName);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                Utlity.Log("File Already Exists," + outputXLfileName, logFilePath);                
                //xlWorkbook = xlApp.Workbooks.Open(outputXLfileName);
                try
                {
                    xlWorkbook = workbooks.Open(outputXLfileName);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = workbooks.Open(outputXLfileName, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Utlity.Log("File Does Not Exist, Use TVS to Generate Template Excel,", logFilePath);
                return false;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return false;
            }


            Utlity.Log("Updating Part Sheets", logFilePath);

            Dictionary<String, List<Variable>> variableDictionaryDetails = Utlity.BuildVariableDictionary(variablesList, logFilePath);
            foreach (String occurenceName in variableDictionaryDetails.Keys)
            {
                List<Variable> variablesList1 = null;
                variableDictionaryDetails.TryGetValue(occurenceName, out variablesList1);
                if (variablesList1.Count != 0)
                {
                    Utlity.Log(occurenceName + ":::" + variablesList1.Count.ToString(), logFilePath);
                    try
                    {
                        ProcessOccurrence(xlWorkbook, occurenceName, variablesList1, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                        Utlity.Log(ex.StackTrace, logFilePath);
                    }
                }
            }

            xlWorkbook.Close(true);

            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

            Utlity.Log("Ending SaveDeltaToXL: " + outputXLfileName, logFilePath);
            return true;

        }

        private static void ProcessOccurrence(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String ocurrenceName, List<Variable> variableArr, String logFilePath)
        {

            bool sheetHide = false;
            Sheets sheets = xlWorkbook.Sheets;
            Microsoft.Office.Interop.Excel._Worksheet sheet = null;
            try
            {
                sheet = xlWorkbook.Sheets[ocurrenceName];
            }
            catch (Exception ex)
            {
                // Could Not Find a Sheet. Need to ADD New One.
                Utlity.Log("WriteDeltaSheet: " + ex.Message, logFilePath);
                sheet = null;
            }

            if (sheet == null)
            {
                sheet = sheets.Add();
                try
                {
                    sheet.Name = ocurrenceName;
                }
                catch (Exception ex)
                {
                    Utlity.Log("WriteDeltaSheet: " + ex.Message, logFilePath);
                    Utlity.Log("WriteDeltaSheet: " + ex.StackTrace, logFilePath);
                    return;
                }
                WriteNewSheet(sheet, variableArr, sheetHide, logFilePath);
            }
            else
            {
                UpdateDeltaInOldSheet(xlWorkbook, sheet, variableArr, sheetHide, logFilePath);
                xlWorkbook.Save();

            }

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //Marshal.ReleaseComObject(xlWorkbook);
            //xlWorkbook = null;


        }

        private static void UpdateDeltaInOldSheet(Microsoft.Office.Interop.Excel._Workbook xlWorkbook, Microsoft.Office.Interop.Excel._Worksheet sheet, List<Variable> variableArrFromEdge, bool sheetHide,
                    String logFilePath)
        {
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            List<Variable>VariableArrFromExcel = readOccurenceVariablesFromTemplateExcel(xlWorkSheet, logFilePath);
            if (VariableArrFromExcel == null || VariableArrFromExcel.Count == 0)
            {
                Utlity.Log("VariableArrFromExcel is Zero.." , logFilePath);
                return;
            }
            Utlity.Log("VariableArrFromExcel: " + VariableArrFromExcel.Count, logFilePath);
            List<String> VariableNamesToBeDeleted = DeterMineRowsToBeDeleted(VariableArrFromExcel, variableArrFromEdge, logFilePath);

            List<String> VariableNamesToBeAdded = DeterMineVariablesToBeAdded(VariableArrFromExcel, variableArrFromEdge, logFilePath);
            if (VariableNamesToBeAdded != null && VariableNamesToBeAdded.Count > 0)
            {
                AddNewVariablesToPartSheet(sheet, VariableNamesToBeAdded, variableArrFromEdge, logFilePath);
            }

            Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkSheet = null;

            xlWorkbook.Save();

        }

        private static void AddNewVariablesToPartSheet(_Worksheet sheet, List<string> VariableNamesToBeAdded, List<Variable> variableArrFromEdge, string logFilePath)
        {
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;

            for (int i = xlRange.Rows.Count + 1; i < (xlRange.Rows.Count + VariableNamesToBeAdded.Count); i++)
            {

            }



            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

            Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkSheet = null;

        
          
        }

        private static void WriteNewSheet(Microsoft.Office.Interop.Excel._Worksheet sheet, List<Variable> variableArr, bool sheetHide,
            String logFilePath)
        {
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;

            xlRange.Cells[1, 1].Value2 = "Name";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "A", logFilePath);
            xlRange.Cells[1, 2].Value2 = "systemName";
            //Utlity.HideExcelColumn(xlWorkSheet, "B", logFilePath);            
            xlRange.Cells[1, 3].Value2 = "value";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "C", logFilePath);
            xlRange.Cells[1, 4].Value2 = "unit";
            //Utlity.HideExcelColumn(xlWorkSheet, "D", logFilePath);            
            xlRange.Cells[1, 5].Value2 = "rangeLow";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "E", logFilePath);
            xlRange.Cells[1, 6].Value2 = "rangeHigh";
            // Utlity.AutoFitExcelColumn(xlWorkSheet, "F", logFilePath);
            xlRange.Cells[1, 7].Value2 = "rangeCondition";
            //Utlity.HideExcelColumn(xlWorkSheet, "G", logFilePath);
            xlRange.Cells[1, 8].Value2 = "Formula";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "H", logFilePath);
            xlRange.Cells[1, 9].Value2 = "PartName";
            xlRange.Cells[1, 10].Value2 = "AddPartToTemplate";
            xlRange.Cells[1, 11].Value2 = "AddVarToTemplate";
            xlRange.Cells[1, 12].Value2 = "LOV";
            xlRange.Cells[1, 13].Value2 = "variableType";
            xlRange.Cells[1, 14].Value2 = "DefaultValue";
            xlRange.Cells[1, 15].Value2 = "UnitType";


            for (int i = 0; i < variableArr.Count; i++)
            {
                xlRange.Cells[i + 2, 1].Value2 = variableArr[i].name;
                xlRange.Cells[i + 2, 2].Value2 = variableArr[i].systemName;
                xlRange.Cells[i + 2, 3].Value2 = variableArr[i].value;
                xlRange.Cells[i + 2, 4].Value2 = variableArr[i].unit;
                xlRange.Cells[i + 2, 5].Value2 = variableArr[i].rangeLow;
                xlRange.Cells[i + 2, 6].Value2 = variableArr[i].rangeHigh;
                xlRange.Cells[i + 2, 7].Value2 = variableArr[i].rangeCondition;
                if (variableArr[i].Formula != null && variableArr[i].Formula.StartsWith("=") == true)
                {
                    String Formula = variableArr[i].Formula;
                    xlRange.Cells[i + 2, 8].Value2 = "'" + Formula;

                    // BOLD if cell has FORMULA
                    //Microsoft.Office.Interop.Excel.Range NewRange = sheet.Cells[i + 2, 8];
                    //NewRange.EntireRow.Select();
                    //NewRange.Font.Bold = true;
                    //Marshal.ReleaseComObject(NewRange);
                    //NewRange = null;
                }
                else
                {
                    if (variableArr[i].Formula != null)
                    {
                        xlRange.Cells[i + 2, 8].Value2 = variableArr[i].Formula;
                        // BOLD if cell has FORMULA
                        //Microsoft.Office.Interop.Excel.Range NewRange = sheet.Cells[i + 2, 8];
                        //NewRange.EntireRow.Select();
                        //NewRange.Font.Bold = true;
                        //Marshal.ReleaseComObject(NewRange);
                        //NewRange = null;
                    }
                    else
                    {

                        xlRange.Cells[i + 2, 8].Value2 = variableArr[i].Formula;
                    }
                }
                xlRange.Cells[i + 2, 9].Value2 = variableArr[i].PartName;
                xlRange.Cells[i + 2, 10].Value2 = variableArr[i].AddPartToTemplate;
                if (variableArr[i].AddPartToTemplate != null)
                {
                    if (variableArr[i].AddPartToTemplate.Equals("N", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        sheetHide = true;
                    }
                    else
                    {
                        sheetHide = false;
                    }
                }
                else
                {
                    //AddPartToTemplate
                    xlRange.Cells[i + 2, 10].Value2 = "N";
                    sheetHide = false;
                }
                xlRange.Cells[i + 2, 11].Value2 = variableArr[i].AddVarToTemplate.ToString();
                // 5 SEPT - Adding a Quote to Make excel remember it as a String.
                xlRange.Cells[i + 2, 12].Value2 = "'" + variableArr[i].LOV;
                if (xlRange.Cells[i + 2, 12].Value2 != null && variableArr[i].LOV != null)
                {
                    // ADD the LOV to value.
                    ExcelUtils.AddLOVtoColumn(xlWorkSheet, 3, i + 2, 12, i + 2, logFilePath);
                }
                xlRange.Cells[i + 2, 13].Value2 = variableArr[i].variableType;
                xlRange.Cells[i + 2, 14].Value2 = variableArr[i].DefaultValue;
                xlRange.Cells[i + 2, 15].Value2 = variableArr[i].UnitType;
            }


            //xlRange.Select();
            //xlRange.Columns.AutoFit();
            //xlRange.Rows.AutoFit();
            //xlRange.Rows.AutoFit();
            //xlRange.Columns.AutoFit();

            xlWorkSheet.UsedRange.EntireColumn.AutoFit();
            xlWorkSheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            FormatCondition format = xlWorkSheet.UsedRange.Rows.FormatConditions.Add(Type: XlFormatConditionType.xlExpression, Operator: Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlEqual, Formula1: "=\"MOD(ROW(),2) = 0\"");
            format.Interior.Color = XlRgbColor.rgbLightBlue;

            //Microsoft.Office.Interop.Excel.FormatConditions fcs = xlRange.FormatConditions;
            //Microsoft.Office.Interop.Excel.FormatCondition fc = (Microsoft.Office.Interop.Excel.FormatCondition)fcs.Add
            //    (Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression, Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlEqual, "=\"MOD(ROW(),2) = 0\"");
            //Microsoft.Office.Interop.Excel.Interior interior = fc.Interior;
            //interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //Marshal.ReleaseComObject(fcs);
            //Marshal.ReleaseComObject(fc);
            //Marshal.ReleaseComObject(interior);
            //interior = null;
            //fc = null;
            //fcs = null;

            Utlity.HideExcelColumn(xlWorkSheet, "D", logFilePath);
            //Utlity.HideExcelColumn(xlWorkSheet, "E", logFilePath);
            //Utlity.HideExcelColumn(xlWorkSheet, "F", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "G", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "H", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "L", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "M", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "N", logFilePath); // DefaultValue
            Utlity.HideExcelColumn(xlWorkSheet, "O", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "B", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "I", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "J", logFilePath);
            Utlity.HideExcelColumn(xlWorkSheet, "K", logFilePath);

            try
            {
                ExcelUtils.HideRows(xlWorkSheet, xlRange, variableArr, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
                Utlity.Log(ex.StackTrace, logFilePath);
            }

            // Column and Row Hide should be done before Hiding the Sheet -- IMP
            if (sheetHide == true)
            {
                xlWorkSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
            }
            else if (sheetHide == false)
            {
                xlWorkSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;
            }

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

            Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkSheet = null;
        }


        private static List<String> DeterMineRowsToBeDeleted(List<Variable> VariableArrFromExcel,
            List<Variable> variableArrFromEdge, String logFilePath)
        {
            List<String> VariableNamesToBeDeleted = 
    VariableArrFromExcel.Select(c => c.name).Except(variableArrFromEdge.Select(d => d.name)).ToList();
            Utlity.Log("VariableNamesToBeDeleted: " + VariableNamesToBeDeleted.Count,logFilePath);
            return VariableNamesToBeDeleted;


        }

        private static List<String> DeterMineVariablesToBeAdded(List<Variable> VariableArrFromExcel,
            List<Variable> variableArrFromEdge, String logFilePath)
        {

            List<String> VariableNamesToBeAdded =
    variableArrFromEdge.Select(c => c.name).Except(VariableArrFromExcel.Select(d => d.name)).ToList();

            Utlity.Log("VariableNamesToBeAdded: " + VariableNamesToBeAdded.Count, logFilePath);

            return VariableNamesToBeAdded;
        }


        public static List<Variable> readOccurenceVariablesFromTemplateExcel(Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet, String logFilePath)
        {

            List<Variable> ExcelSheetVariableList = new List<Variable>();
            Utlity.Log("----------------------------------------------------------", logFilePath);

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    if (i == 1)
                        continue;
                    try
                    {
                        //Utlity.Log("Iteration: " + i.ToString(), logFilePath);

                        Variable varr = new Variable();
                        if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                        {
                            try
                            {

                                varr.name = xlRange.Cells[i, 1].Value2;

                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("name" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            Utlity.Log("Skip Row Number : " + i, logFilePath);
                            continue;
                        }
                        if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2] != null)
                        {
                            try
                            {
                                varr.systemName = xlRange.Cells[i, 2].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("systemName" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        {
                            try
                            {
                                varr.value = xlRange.Cells[i, 3].Value2;
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("value" + ex.Message, logFilePath);
                                double value = xlRange.Cells[i, 3].Value2;
                                varr.value = value.ToString("0.######");
                            }
                        }
                        if (xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4] != null)
                        {
                            try
                            {
                                varr.unit = xlRange.Cells[i, 4].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("unit" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                        {
                            try
                            {
                                varr.rangeLow = xlRange.Cells[i, 5].Value2;
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("rangeLow" + ex.Message, logFilePath);
                                double value = xlRange.Cells[i, 5].Value2;
                                varr.rangeLow = value.ToString("0.######");

                            }
                        }
                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {
                            try
                            {
                                varr.rangeHigh = xlRange.Cells[i, 6].Value2;
                            }
                            catch (Exception ex)
                            {
                                double value = xlRange.Cells[i, 6].Value2;
                                varr.rangeHigh = value.ToString("0.######");
                                //Utlity.Log("rangeHigh" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        {
                            try
                            {
                                int result = 0;
                                Double rangeCondition = xlRange.Cells[i, 7].Value2;
                                result = (int)rangeCondition;
                                varr.rangeCondition = result;

                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Trouble Parsing rangeCondition: " + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                        {
                            try
                            {
                                varr.Formula = xlRange.Cells[i, 8].Value2;
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("Formula: " + ex.Message, logFilePath);
                                int Formula = xlRange.Cells[i, 8].Value2;
                                varr.Formula = Formula.ToString();
                            }
                        }
                        if (xlRange.Cells[i, 9].Value2 != null && xlRange.Cells[i, 9] != null)
                        {
                            try
                            {
                                varr.PartName = xlRange.Cells[i, 9].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("PartName" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 10].Value2 != null && xlRange.Cells[i, 10] != null)
                        {
                            try
                            {
                                varr.AddPartToTemplate = xlRange.Cells[i, 10].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AddPartToTemplate" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 11] != null)
                        {
                            try
                            {
                                varr.AddVarToTemplate = (bool)xlRange.Cells[i, 11].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AddVarToTemplate" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 12].Value2 != null && xlRange.Cells[i, 12] != null)
                        {
                            try
                            {
                                varr.LOV = xlRange.Cells[i, 12].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("LOV" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 13].Value2 != null && xlRange.Cells[i, 13] != null)
                        {
                            try
                            {
                                varr.variableType = xlRange.Cells[i, 13].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("variableType" + ex.Message, logFilePath);
                            }
                        }
                        if (xlRange.Cells[i, 14].Value2 != null && xlRange.Cells[i, 14] != null)
                        {
                            try
                            {
                                varr.DefaultValue = xlRange.Cells[i, 14].Value2;
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("value" + ex.Message, logFilePath);
                                double DefaultValue = xlRange.Cells[i, 14].Value2;
                                varr.DefaultValue = DefaultValue.ToString("0.######");
                            }
                        }

                        if (xlRange.Cells[i, 15].Value2 != null && xlRange.Cells[i, 15] != null)
                        {
                            try
                            {
                                varr.UnitType = xlRange.Cells[i, 15].Value2;
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("UnitType" + ex.Message, logFilePath);

                            }
                        }
                        ExcelSheetVariableList.Add(varr);
                       
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                    }


                }
                Marshal.ReleaseComObject(xlRange);
                xlRange = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            return ExcelSheetVariableList;

        }

        
    }
}
