using ExcelSyncTC.controller;
using ExcelSyncTC.model;
using ExcelSyncTC.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelSyncTC.opInterfaces
{
    class ExcelInterface
    {

        private static void AddLOVtoColumn(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, int columnIndex, int RowIndex, int lovColumnIndex, int lovRowIndex, String logFilePath)
        {
            String columnLetter = ColumnIndexToColumnLetter(columnIndex);
            var cell1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Range[columnLetter + RowIndex.ToString(), columnLetter + RowIndex.ToString()];
            if (cell1 == null)
            {
                //utils.Utlity.Log("AddLOVtoColumn: " + "cell1 is Empty", logFilePath);                    
                return;
            }
            columnLetter = ColumnIndexToColumnLetter(lovColumnIndex);
            var cell2 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Range[columnLetter + lovRowIndex.ToString(), columnLetter + lovRowIndex.ToString()];
            if (cell2 == null)
            {
                //utils.Utlity.Log("AddLOVtoColumn: " + "cell2 is Empty", logFilePath); 
                return;
            }
            String LovValues = "";
            try
            {
                LovValues = (String)cell2.Value2;
            }
            catch (Exception ex)
            {
                //utils.Utlity.Log("AddLOVtoColumn: " + "LovValues Exception: " + ex.Message, logFilePath);
                return;
            }

            if (LovValues == null || LovValues.Equals("") == true)
            {
                //utils.Utlity.Log("AddLOVtoColumn: " + "LovValues is Empty", logFilePath);
                return;
            }
            Utlity.Log(LovValues.ToString(), logFilePath);
            //LovValues = LovValues.Trim('\'');
            //Utlity.Log(LovValues.ToString(), logFilePath);
            String[] lovArray = LovValues.Split(',');
            //foreach (String s in lovArray)
            //{
            //    Utlity.Log(s, logFilePath);
            //}

            var FlatList = string.Join(",", lovArray.ToArray());

            cell1.Validation.Delete();

            try
            {
                cell1.Validation.Add(
                   XlDVType.xlValidateList,
                   AlertStyle: XlDVAlertStyle.xlValidAlertInformation,
                   Operator: XlFormatConditionOperator.xlBetween,
                   Formula1: FlatList
                   );
                cell1.Validation.IgnoreBlank = true;
                cell1.Validation.InCellDropdown = true;
            }
            catch (Exception ex)
            {
                Utlity.Log("Validation Add: " + ex.Message, logFilePath);
                return;
            }

            //Console.WriteLine(cell1.Validation.Formula1);
            xlWorksheet.Activate();
        }

        private static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        public static bool SaveDeltaToXL(Microsoft.Office.Interop.Excel._Application xlApp,String outputXLfileName,
           String logFilePath)
        {
            Utlity.Log("Inside SaveDeltaToXL", logFilePath);
           

            Dictionary<String, List<Variable>> variableDictionary = SolidEdgeData1.getVariablesDictionaryDetails();
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utlity.Log("variableDictionary is NULL,", logFilePath);
                return false;
            }

            
            xlApp.DisplayAlerts = false;

            
            FileInfo f = new FileInfo(outputXLfileName);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                Utlity.Log("File Already Exists,", logFilePath);
                xlWorkbook = xlApp.ActiveWorkbook;
            }
            else
            {
             
                return false ;
            }
            if (xlWorkbook == null)
            {
              
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return false;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;
            Utlity.Log("Reading WorkSheets..,", logFilePath);
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

                 // 29 - SEPT, Ignore Sheets that Are not changed by the User/External Link
                 if (Ribbon1.dialog.syncAllCb.Checked == false) //Change requested by LTC to sync all sheets 25-Dec-2019
                 {
                     if (Utlity.ModSheetsInSession != null && Utlity.ModSheetsInSession.Contains(sheet.Name) == false)
                     {
                         Utlity.Log(sheet.Name + "::::" + " Is not Modified By User, Skipping It", logFilePath);
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


           

            if (xlApp != null) xlApp.DisplayAlerts = true;
            return true;
        }

        // 12- SEPT Writes the Modified Value from variableArr List to the Excel Sheet.
        // SyncTE would update the Value and Reevaluate the Expressions in SE and Write back the Data into XL.
        private static void WriteSheet(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel.Worksheet sheet,String ocurrenceName, List<Variable> variableArr, String logFilePath)
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
               if (xlRange.Cells[i, 1].Value2 != null || xlRange.Cells[i, 1] != null)
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

        


    }
}
