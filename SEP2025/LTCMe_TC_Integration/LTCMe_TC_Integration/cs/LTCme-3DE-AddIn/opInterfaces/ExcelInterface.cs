/**************************************************************************
 *  3DE International FZE, 2018
 *  Murali - 02-OCT-18 - Changes for SuppressionEnabled Property
 *  ***********************************************************************
 */

using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DemoAddInTC.opInterfaces
{
    class ExcelInterface
    {

        


        public static bool SaveToXL(String outputXLfileName, List<Variable>variablesList,
            Dictionary<String,bool> partEnablementDictionary,String logFilePath,String option,List<FeatureLine>fsList)
        {
            Utlity.Log("Inside SaveToXL", logFilePath);
            if (partEnablementDictionary == null)
            {
                Utlity.Log("partEnablementDictionary is NULL,", logFilePath);
                return false;
            }
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            //xlApp.WindowState = XlWindowState.xlNormal;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            FileInfo f = new FileInfo(outputXLfileName);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                Utlity.Log("File Already Exists,", logFilePath);
                return false;
                //xlWorkbook = xlApp.Workbooks.Open(outputXLfileName);
            }
            else
            {
                xlWorkbook = xlApp.Workbooks.Add(Type.Missing);
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return false;
            }

            //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.ActiveSheet;
            Utlity.Log("Creating Part Sheets", logFilePath);
            //Dictionary<String,List<Variable>> variableDictionaryDetails = SolidEdgeData.getVariablesDictionaryDetails();
            Dictionary<String, List<Variable>> variableDictionaryDetails = Utlity.BuildVariableDictionary(variablesList,logFilePath);
            foreach (String occurenceName in variableDictionaryDetails.Keys) {
                List<Variable> variablesList1 = null;
                variableDictionaryDetails.TryGetValue(occurenceName,out variablesList1 );
                if (variablesList1.Count != 0)
                {
                    Utlity.Log(occurenceName + ":::" + variablesList1.Count.ToString(), logFilePath);
                    try
                    {
                        //WriteSheet(xlWorkbook, occurenceName, variablesList1, logFilePath);
                        WriteSheetFast(xlWorkbook, occurenceName, variablesList1, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                        Utlity.Log(ex.StackTrace, logFilePath);
                    }
                }
            }

            // Need to UPDATE Part Enabled/NOT in this Sheet
            Utlity.Log("Creating MASTER ASSEMBLY Sheet", logFilePath);
            List<BOMLine> BomLineList = null;
            String assemblyFilePath = Path.ChangeExtension(outputXLfileName,".asm");
            if (option.Equals("TVS") == true)
            {
                BomLineList = SolidEdgeData1.getBomLinesList();
            }
            else
            {
                BomLineList = ExcelData.getBomLineList();
            }
            if (BomLineList.Count == 0)
            {
                Utlity.Log("BomLineList Count is Zero", logFilePath);                
            }
            else
            {
                try
                {
                    WriteBOMStructure(xlWorkbook, "MASTER ASSEMBLY", BomLineList, partEnablementDictionary, logFilePath);
                }
                catch (Exception ex)
                {
                    Utlity.Log("WriteBOMStructure: " + ex.Message, logFilePath);
                }
            }

            // Write Feature Suppression data to FEATURES sheet.
            Utlity.Log("Writing Feature Information...", logFilePath);
            if (fsList.Count == 0 || fsList == null)
            {
                Utlity.Log("fsList Count is Zero", logFilePath);
            }
            else
            {
                WriteFeatureData(xlWorkbook, "FEATURES", fsList, logFilePath);
            }


            Utlity.Log("Deleting Sheets - Foglio & Sheet", logFilePath);
            List<Microsoft.Office.Interop.Excel._Worksheet> sheetList = new List<_Worksheet>();
            try
            {
                foreach (Microsoft.Office.Interop.Excel._Worksheet sheet in xlWorkbook.Sheets)
                {
                    if (sheet.Name.StartsWith("Foglio", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        sheetList.Add(sheet);
                        //sheet.Delete();
                    }

                    if (sheet.Name.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        sheetList.Add(sheet);
                        //sheet.Delete();
                    }
                    //Marshal.ReleaseComObject(sheet);
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("Deleting Sheets - Exception: " + ex.Message, logFilePath);
                Utlity.Log("Deleting Sheets - Exception: " + ex.StackTrace, logFilePath);
            }
            try
            {
                for (int i = 0; i < sheetList.Count; i++)
                {
                    sheetList[i].Delete();
                    sheetList[i] = null;
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("Deleting Sheets - Exception: " + ex.Message, logFilePath);
                Utlity.Log("Deleting Sheets - Exception: " + ex.StackTrace, logFilePath);
            }

            sheetList.Clear();
            //xlWorksheet.Delete(); // I dont need Sheet1
            //release com objects to fully kill excel process from running in the background            
            //Marshal.ReleaseComObject(xlWorksheet);

            xlApp.Visible = false;
            xlApp.UserControl = false;
            if (f.Exists == false)
            {
                Utlity.Log(outputXLfileName + " Being Saved...",logFilePath);
                xlWorkbook.SaveAs(outputXLfileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
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
            variablesList.Clear();
            return true;

        }
    private static void WriteSheet(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String ocurrenceName, List<Variable> variableArr,String logFilePath)
        {
            
            bool sheetHide = false;
            Sheets sheets = xlWorkbook.Sheets;
            var sheet = sheets.Add();
            //var sheet = xlWorkbook.Sheets.Add();
            try
            {
                sheet.Name = ocurrenceName;
            }
            catch (Exception ex)
            {
                Utlity.Log("WriteSheet: " + ex.Message, logFilePath);
                Utlity.Log("WriteSheet: " + ex.StackTrace, logFilePath);
                return;
            }
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
            xlRange.Cells[1,11].Value2 = "AddVarToTemplate";           
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
                if (xlRange.Cells[i + 2, 12].Value2 != null && variableArr[i].LOV !=null )
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
                ExcelUtils.HideRows(xlWorkSheet, xlRange,variableArr, logFilePath);
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
            Marshal.ReleaseComObject(sheets);
            sheets = null;
            Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkSheet = null;

        }


    private static void WriteSheetFast(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String ocurrenceName, List<Variable> variableArr, String logFilePath)
    {

        bool sheetHide = false;
        Sheets sheets = xlWorkbook.Sheets;
        var sheet = sheets.Add();
        //var sheet = xlWorkbook.Sheets.Add();
        try
        {
            sheet.Name = ocurrenceName;
        }
        catch (Exception ex)
        {
            Utlity.Log("WriteSheet: " + ex.Message, logFilePath);
            Utlity.Log("WriteSheet: " + ex.StackTrace, logFilePath);
            return;
        }
        Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
        xlWorkSheet.Activate();

        Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;

        // Create the array.
        int NumOfColumns = 15;
        int NumOfRows = variableArr.Count;
        // +1 for Header
        object[,] myArray = new object[NumOfRows+1, NumOfColumns];

        //Utlity.Log("WriteSheet: " + "Writing the Headers..", logFilePath);

        myArray[0, 0] = "Name";
        //Utlity.AutoFitExcelColumn(xlWorkSheet, "A", logFilePath);
        myArray[0, 1]= "systemName";
        //Utlity.HideExcelColumn(xlWorkSheet, "B", logFilePath);            
        myArray[0, 2] = "value";
        //Utlity.AutoFitExcelColumn(xlWorkSheet, "C", logFilePath);
        myArray[0, 3] = "unit";
        //Utlity.HideExcelColumn(xlWorkSheet, "D", logFilePath);            
        myArray[0, 4] = "rangeLow";
        //Utlity.AutoFitExcelColumn(xlWorkSheet, "E", logFilePath);
        myArray[0, 5] = "rangeHigh";
        // Utlity.AutoFitExcelColumn(xlWorkSheet, "F", logFilePath);
        myArray[0, 6] = "rangeCondition";
        //Utlity.HideExcelColumn(xlWorkSheet, "G", logFilePath);
        myArray[0, 7] = "Formula";
        //Utlity.AutoFitExcelColumn(xlWorkSheet, "H", logFilePath);
        myArray[0, 8] = "PartName";
        myArray[0, 9] = "AddPartToTemplate";
        myArray[0, 10] = "AddVarToTemplate";
        myArray[0, 11] = "LOV";
        myArray[0, 12] = "variableType";
        myArray[0, 13] = "DefaultValue";
        myArray[0, 14] = "UnitType";

        //Utlity.Log("WriteSheet: " + "Writing the Rows..", logFilePath);
        for (int i = 0; i < variableArr.Count; i++)
        {
            // Start with 1st Row, 0th Row is the Header
            myArray[i + 1, 0] = variableArr[i].name;
            Utlity.Log("variableArr[i].name " + variableArr[i].name, logFilePath);
            myArray[i + 1, 1] = variableArr[i].systemName;
            Utlity.Log("variableArr[i].systemName " + variableArr[i].systemName, logFilePath);
            myArray[i + 1,2] = variableArr[i].value;
            Utlity.Log("variableArr[i].value " + variableArr[i].value, logFilePath);
            myArray[i + 1, 3] =variableArr[i].unit;
            myArray[i + 1, 4] = variableArr[i].rangeLow;
            myArray[i + 1, 5] = variableArr[i].rangeHigh;
            myArray[i + 1, 6] = variableArr[i].rangeCondition;

            if (variableArr[i].Formula != null && variableArr[i].Formula.StartsWith("=") == true)
            {
                String Formula = variableArr[i].Formula;
                myArray[i + 1, 7] = "'" + Formula;
            }
            else
            {
                if (variableArr[i].Formula != null)
                {
                    myArray[i + 1, 7] = variableArr[i].Formula;
                }
                else
                {
                    myArray[i + 1, 7] = variableArr[i].Formula;
                }
            }

            myArray[i+1,8] = variableArr[i].PartName;
            myArray[i + 1, 9] = variableArr[i].AddPartToTemplate;
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
                myArray[i + 1, 9] = "N";
                sheetHide = false;
            }

            myArray[i + 1, 10] = variableArr[i].AddVarToTemplate.ToString();
            // 5 SEPT - Adding a Quote to Make excel remember it as a String.
            myArray[i + 1, 11] = "'" + variableArr[i].LOV;



            myArray[i + 1, 12] = variableArr[i].variableType;
            myArray[i + 1, 13] = variableArr[i].DefaultValue;
            myArray[i + 1, 14] = variableArr[i].UnitType;

        }

        //Utlity.Log("WriteSheet: " + "Setting the Range At One Shot..", logFilePath);
        // Create a Range of the correct size:
        int rows = myArray.GetLength(0);
        int columns = myArray.GetLength(1);
        //utils.Utlity.Log("rows: " + rows, logFilePath);
        //utils.Utlity.Log("columns: " + columns, logFilePath);
        Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.get_Range("A1", Type.Missing);
        range = range.get_Resize(rows, columns);
        // Assign the Array to the Range in one shot:
        range.set_Value(Type.Missing, myArray);
        Marshal.ReleaseComObject(range);
        range = null;

       // Utlity.Log("WriteSheet: " + "Adding LOV..", logFilePath);
        for (int i = 0; i < variableArr.Count; i++)
        {
            if (variableArr[i].LOV != null)
            {
                // ADD the LOV to value.
                ExcelUtils.AddLOVtoColumn(xlWorkSheet, 3, i + 2, 12, i + 2, logFilePath);
            }
        }
        //Utlity.Log("WriteSheet: " + "Autofit..", logFilePath);
        xlWorkSheet.UsedRange.EntireColumn.AutoFit();
        xlWorkSheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

        FormatCondition format = xlWorkSheet.UsedRange.Rows.FormatConditions.Add(Type: XlFormatConditionType.xlExpression, Operator: Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlEqual, Formula1: "=\"MOD(ROW(),2) = 0\"");
        format.Interior.Color = XlRgbColor.rgbLightBlue;


        //Utlity.Log("WriteSheet: " + "HideExcelColumn..", logFilePath);
        Utlity.HideExcelColumn(xlWorkSheet, "D", logFilePath);
        
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
           // Utlity.Log("WriteSheet: " + "HideRows..", logFilePath);
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
        Marshal.ReleaseComObject(sheets);
        sheets = null;
        Marshal.ReleaseComObject(xlWorkSheet);
        xlWorkSheet = null;

    }

    


private static void WriteBOMStructure(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String assemblyName, List<BOMLine> BOMLineArr, Dictionary<String,bool> partEnablementDictionary,String logFilePath)
        {
            Sheets sheets = xlWorkbook.Sheets;
            var sheet = sheets.Add();
            sheet.Name = assemblyName;
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            
            xlRange.Cells[1, 1].Value2 = "FullName";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "A", logFilePath);
            xlRange.Cells[1, 2].Value2 = "level";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "B", logFilePath);
            xlRange.Cells[1, 3].Value2 = "AbsolutePath";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "C", logFilePath);
            xlRange.Cells[1, 4].Value2 = "DocNum";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "D", logFilePath);
            xlRange.Cells[1, 5].Value2 = "Revision";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "E", logFilePath);
            xlRange.Cells[1, 6].Value2 = "Status";
            //Utlity.AutoFitExcelColumn(xlWorkSheet, "F", logFilePath);            
            xlRange.Cells[1, 7].Value2 = "FullName";

            for (int i = 1; i <= BOMLineArr.Count; i++)
            {
                String component = BOMLineArr[i - 1].FullName;
                if (component != null && component.Equals("") == false) {
                    
                    
                    xlRange.Cells[i + 1, 1].Value2 = "'" + Path.GetFileNameWithoutExtension(component);
                }
                //xlRange.Cells[i + 1, 1].Value2 = BOMLineArr[i - 1].FullName;

                xlRange.Cells[i + 1, 2].Value2 = BOMLineArr[i - 1].level;
                xlRange.Cells[i + 1, 3].Value2 = BOMLineArr[i - 1].AbsolutePath;
                xlRange.Columns[4].NumberFormat = "00000000";
                xlRange.Cells[i + 1, 4].Value2 = "'" + BOMLineArr[i - 1].DocNum;
                xlRange.Cells[i + 1, 5].Value2 = BOMLineArr[i - 1].Revision;
                bool value = false;
                String compWithExtn = Path.GetFileName(BOMLineArr[i - 1].AbsolutePath);
                if (compWithExtn != null && compWithExtn.Equals("") == false)
                {
                    //Utlity.Log(compWithExtn, logFilePath);
                    partEnablementDictionary.TryGetValue(compWithExtn, out value);
                    if (value == true)
                    {
                        //xlRange.Cells[i + 1, 6].Value2 = BOMLineArr[i - 1].Status;
                        xlRange.Cells[i + 1, 6].Value2 = "INCLUDED";
                    }
                    else
                    {
                        xlRange.Cells[i + 1, 6].Value2 = "EXCLUDED";
                    }
                }
                else
                {
                    xlRange.Cells[i + 1, 6].Value2 = "EXCLUDED";
                }
                xlRange.Cells[i + 1, 7].Value2 = BOMLineArr[i - 1].FullName;
            }

            xlWorkSheet.UsedRange.EntireColumn.AutoFit();
            xlWorkSheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            try
            {
                FormatCondition format = xlWorkSheet.UsedRange.Rows.FormatConditions.Add(XlFormatConditionType.xlExpression, Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2) = 0");
                format.Interior.Color = XlRgbColor.rgbLightBlue;
            }
            catch (Exception ex)
            {
                Utlity.Log("WriteBOMStructure: Formatting Exception: " + ex.Message, logFilePath);
            }

            Utlity.HideExcelColumn(xlWorkSheet, "G", logFilePath);

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;
            Marshal.ReleaseComObject(sheets);
            sheets = null;
            Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkSheet = null;

        }


private static void WriteFeatureData(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String assemblyName, List<FeatureLine> featureLineList, String logFilePath)
{
    Sheets sheets = xlWorkbook.Sheets;
    var sheet = sheets.Add();
    sheet.Name = assemblyName;
    Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
    xlWorkSheet.Activate();

    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;    
    
    xlRange.Cells[1, 1].Value2 = "PartName";
    //Utlity.AutoFitExcelColumn(xlWorkSheet, "A", logFilePath);
    xlRange.Cells[1, 2].Value2 = "FeatureName";
    //Utlity.AutoFitExcelColumn(xlWorkSheet, "B", logFilePath);
    xlRange.Cells[1, 3].Value2 = "SystemName";
    //Utlity.AutoFitExcelColumn(xlWorkSheet, "C", logFilePath);
    xlRange.Cells[1, 4].Value2 = "Formula";
    //Utlity.AutoFitExcelColumn(xlWorkSheet, "D", logFilePath);
    xlRange.Cells[1, 5].Value2 = "IsFeatureEnabled";
    //Utlity.AutoFitExcelColumn(xlWorkSheet, "E", logFilePath);
    xlRange.Cells[1, 6].Value2 = "PartName";
    xlRange.Cells[1, 7].Value2 = "EdgeBarName";
    xlRange.Cells[1, 8].Value2 = "SuppressionEnabled"; // 02 - OCT, Added Based on LTC Request.


    for (int i = 1; i <= featureLineList.Count; i++)
    {
        String PartName = featureLineList[i - 1].PartName;
        if (PartName != null && PartName.Equals("") == false)
        {
            xlRange.Cells[i + 1, 1].Value2 = PartName;
        }        
        xlRange.Cells[i + 1, 2].Value2 = featureLineList[i - 1].FeatureName;
        xlRange.Cells[i + 1, 3].Value2 = featureLineList[i - 1].SystemName;
        xlRange.Cells[i + 1, 4].Value2 = featureLineList[i - 1].Formula;
        xlRange.Cells[i + 1, 5].Value2 = featureLineList[i - 1].IsFeatureEnabled;
        if (PartName != null && PartName.Equals("") == false)
        {
            xlRange.Cells[i + 1, 6].Value2 = PartName;
        }
        xlRange.Cells[i + 1, 7].Value2 = featureLineList[i-1].EdgeBarName;
        xlRange.Cells[i + 1, 8].Value2 = featureLineList[i - 1].SuppressionEnabled;
    }

    MergeColumn(xlWorkbook,xlWorkSheet, xlRange, logFilePath);
    xlWorkSheet.UsedRange.EntireColumn.AutoFit();

    xlWorkSheet.UsedRange.EntireColumn.AutoFit();
    xlWorkSheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

    //xlWorkSheet.UsedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
    //FormatCondition format = xlWorkSheet.UsedRange.Rows.FormatConditions.Add(XlFormatConditionType.xlExpression, Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2) = 0");
    //format.Interior.Color = XlRgbColor.rgbLightBlue;

    Utlity.HideExcelColumn(xlWorkSheet, "E", logFilePath);
    Utlity.HideExcelColumn(xlWorkSheet, "F", logFilePath);
    Utlity.HideExcelColumn(xlWorkSheet, "G", logFilePath);

    Marshal.ReleaseComObject(xlRange);
    xlRange = null;
    Marshal.ReleaseComObject(sheets);
    sheets = null;
    Marshal.ReleaseComObject(xlWorkSheet);
    xlWorkSheet = null;

}

private static void MergeColumn(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet, Microsoft.Office.Interop.Excel.Range xlRange, String logFilePath)
{

    var style = xlWorkbook.Styles.Add("Merge");
    style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenterAcrossSelection;
    style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
    
    String pName = "";
    int StartRowIndex = 1;
    Microsoft.Office.Interop.Excel.Range FeatureRange = xlWorkSheet.UsedRange;
    //Utlity.Log("MergeColumn: " + FeatureRange.Rows.Count, logFilePath);
    for (int i = 1; i <= FeatureRange.Rows.Count; i++)
    {
        if (i == 1)
            continue;

        if (FeatureRange.Cells[i, 1].Value2 != null && FeatureRange.Cells[i, 1] != null)
        {

            String partName = (String)FeatureRange.Cells[i, 1].Value2;
            //Utlity.Log("MergeColumn: " + partName, logFilePath);
            if (partName == null || partName.Equals("") == true)
            {
                continue;
            }
            if (pName == null || pName.Equals("") == true)
            {
                pName = partName;
                StartRowIndex = i;
                //Utlity.Log("MergeColumn: StartRowIndex: " + StartRowIndex, logFilePath);
            }
            else if (partName.Equals(pName, StringComparison.OrdinalIgnoreCase) == false)
            {
                
                int endRowIndex = i-1;
                //Utlity.Log("MergeColumn: endRowIndex: " + endRowIndex, logFilePath);
                // Merge
                try
                {
                    if (StartRowIndex != endRowIndex)
                    {
                        Microsoft.Office.Interop.Excel.Range s1 = xlWorkSheet.Cells[StartRowIndex, 1];
                        Microsoft.Office.Interop.Excel.Range s2 = xlWorkSheet.Cells[endRowIndex, 1];

                        //Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.get_Range(s1, s2);
                        Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.get_Range("A" + StartRowIndex.ToString(), "A"+ endRowIndex.ToString() );
                        range.Select();
                        range.Merge(false);
                        range.Style = style.Name;
                        Marshal.ReleaseComObject(range);
                        range = null;
                        Marshal.ReleaseComObject(s1);
                        s1 = null;
                        Marshal.ReleaseComObject(s2);
                        s2 = null;
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log("MergeColumn: Exception : " + ex.Message, logFilePath);
                }
                pName = partName;
                StartRowIndex = i;

            }

        }

    }

    Marshal.ReleaseComObject(FeatureRange);
    FeatureRange = null;
}

      // 21 - SEPT - 2018 
      // Rename Sheet -- Add Suffix
public static void RenameSheetNamesForVariableParts(String FolderToPublish, String assemblyFileName, List<String> variablePartsList, String Suffix, String logFilePath)
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

    Utlity.Log("RenameSheetNamesForVariableParts: XLFilePath: " + xlFilePath, logFilePath);
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
    // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
    String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);

    Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
    foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
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

        if (variablePartsList.Contains(sheet.Name) == true)
        {
            //String fileName = System.IO.Path.GetFileName(sheet.Name);
            //if (variablePartsList.Contains(fileName) == true)
            {

                String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(sheet.Name);
                String extn = System.IO.Path.GetExtension(sheet.Name);
                String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                // Change SheetName
                // Update PartName Field in the Sheet
                UpdateSuffixToSheetData(sheet, newPartFileName, logFilePath);

            }            
        }

        Marshal.ReleaseComObject(sheet);
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

    // In case the TopLine is itself a Variable Part, then Need to Change the Name of the Template XL File As Well.
    String AssemFileName = System.IO.Path.GetFileName(assemblyFileName);
    if (variablePartsList.Contains(AssemFileName) == true)
    {

        String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(AssemFileName);         
        String newPartFileName = fileNameWithoutExtn + Suffix + ".xlsx";
        String newPartFileNameFullpath = System.IO.Path.Combine(FolderToPublish, newPartFileName);
        // Change SheetName
        // Update PartName Field in the Sheet
        try
        {
            Utlity.Log(newPartFileNameFullpath, logFilePath);
            File.Move(xlFilePath, newPartFileNameFullpath);
        }
        catch (Exception ex)
        {
            Utlity.Log("Exception in Renaming Template Excel Sheet: " + ex.Message, logFilePath);
        }
    } 
        
}

private static void UpdateSuffixToSheetData(Worksheet sheet, string NewSheetName,String logFilePath)
{
    if (sheet == null)
    {
        return;
    }
    Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
    xlWorkSheet.Activate();

     Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;

          
     //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
     for (int i = 1; i <= xlRange.Rows.Count; i++)
     {
         if (i == 1)
             continue;

         // 9 - PartName is Changed.
         if (xlRange.Cells[i, 9].Value2 != null && xlRange.Cells[i, 9] != null)
         {
             try
             {
                 xlRange.Cells[i, 9].Value2 = NewSheetName;
             }
             catch (Exception ex)
             {
                 Utlity.Log("PartName" + ex.Message, logFilePath);
             }
         }

     }
     try
     {
         // 29 - SEPT - If Sheet Name is more than 31, Renaming of Sheet FAILS.
         sheet.Name = NewSheetName;
     }
     catch (Exception ex)
     {
         Utlity.Log("UpdateSuffixToSheetData: " + ex.Message, logFilePath);
     }
}


    }
}
