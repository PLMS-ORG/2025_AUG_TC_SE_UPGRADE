using DemoAddInTC.model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace DemoAddInTC.utils
{
    class ExcelUtils
    {

        public static void AddLOVtoColumn(Microsoft.Office.Interop.Excel._Worksheet xlWorksheet, int columnIndex, int RowIndex, int lovColumnIndex, int lovRowIndex, String logFilePath)
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

        public static void HideRows(Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet, Microsoft.Office.Interop.Excel.Range xlRange, List<Variable> variableArr, String logFilePath)
        {
            //Utlity.Log("XLRange Count: " + variableArr.Count.ToString(), logFilePath);

            for (int i = 0; i < variableArr.Count; i++)
            {
                bool AddVarToTemplate = (bool)xlRange.Cells[i + 2, 11].Value2;
                //Utlity.Log("AddVarToTemplate: " + AddVarToTemplate, logFilePath);
                //if (AddVarToTemplate != null && AddVarToTemplate.Equals("") == false)
                {
                    if (AddVarToTemplate == false)
                    {
                        Utlity.HideExcelRow(xlWorkSheet, i + 2, logFilePath);
                    }
                }
                //else
                //{
                //    Utlity.HideExcelRow(xlWorkSheet, i + 2, logFilePath);
                //}

            }

        }

        public static void KillExcelFileProcess(String logFilePath)
        {
            try
            {
                var processes = from p in Process.GetProcessesByName("EXCEL")
                                select p;

                foreach (var process in processes)
                {
                    //if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("KillExcelFileProcess: " + ex.Message, logFilePath);
            }
        }
    }
}
