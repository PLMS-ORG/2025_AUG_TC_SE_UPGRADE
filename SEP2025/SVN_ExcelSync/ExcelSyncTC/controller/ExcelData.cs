using ExcelSyncTC.model;
using ExcelSyncTC.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelSyncTC.controller
{
    class ExcelData
    {
        public static List<String> occurenceList = new List<string>();
        // 10 - SEPT - Contains Only Variables from Sheets Which are Visible, Others are Skipped
        public static List<Variable> ALLvariablesList = new List<Variable>();
        // 10 - SEPT - Contains Only Parts Whose Sheets are Visible, Others are Skipped
        public static Dictionary<String, List<Variable>> variableDictionary = new Dictionary<string, List<Variable>>
            ();
        // 10 - SEPT - Contains both Parts that are VISIBLE and INVISIBLE
        public static Dictionary<String, bool> partEnablementDictionary = new Dictionary<string, bool>();
        public static Dictionary<String, String> occurencePathDictionary = new Dictionary<string, string>();

        public ExcelData()
        {
        }

        public static List<Variable> getVariableDetails()
        {
            return ALLvariablesList;
        }

        public static Dictionary<string, bool> getPartEnablementDictionary()
        {
            return partEnablementDictionary;

        }
        public static Dictionary<String, List<Variable>> getVariableDictionary()
        {
            return variableDictionary;
        }

        public static Dictionary<String, String> getOcurrencePathDictionary()
        {
            return occurencePathDictionary;
        }

        public static List<Variable> readOccurenceVariablesFromTemplateExcel(Microsoft.Office.Interop.Excel._Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            ALLvariablesList.Clear();
            partEnablementDictionary.Clear();
            variableDictionary.Clear();
            occurenceList.Clear();

            Utlity.Log("----------------------------------------------------------", logFilePath);            
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return null;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;

            String topLineAssembly = Path.ChangeExtension(xlWorkbook.FullName, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            Utlity.Log("Sheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
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

               
                 List<Variable> variableListForSheet = new List<Variable>();
                Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                try
                {
                    if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        // 10 - SEPT - NOT READING SHEETS WHICH ARE HIDDEN. PERFORMANCE OVER HEAD.
                        partEnablementDictionary.Add(sheet.Name, false);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                    else
                    {
                        partEnablementDictionary.Add(sheet.Name, true);
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }

                Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
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
                                varr.AddVarToTemplate = xlRange.Cells[i, 11].Value2;
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
                        ALLvariablesList.Add(varr);
                        variableListForSheet.Add(varr);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                    }

                    
                }
                Utlity.Log("Releasing Range", logFilePath);
                Marshal.ReleaseComObject(xlRange);
                xlRange = null;

                

                variableDictionary.Add(sheet.Name, variableListForSheet);

                Utlity.Log("Releasing sheet", logFilePath);
                Marshal.ReleaseComObject(sheet);
                //Utlity.Log(sheet.Name + " is Done", logFilePath);  
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Utlity.Log("Releasing Sheets", logFilePath);
            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //release com objects to fully kill excel process from running in the background            
            //Marshal.ReleaseComObject(xlWorksheet);

            //xlApp.Visible = false;
            //xlApp.UserControl = false;           
            Utlity.Log("Releasing workbook", logFilePath);
            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;
            Utlity.Log("----------------------------------------------------------", logFilePath);
            //quit and release
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);
            return ALLvariablesList;

        }

        public static List<Variable> readOccurenceVariablesFromTemplateExcelFast(Microsoft.Office.Interop.Excel._Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            ALLvariablesList.Clear();
            partEnablementDictionary.Clear();
            variableDictionary.Clear();
            occurenceList.Clear();

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return null;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;

            String topLineAssembly = Path.ChangeExtension(xlWorkbook.FullName, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            //Utlity.Log("Sheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in xlWorkbook.Worksheets)
            {
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    //if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    string sheetName = sheet.Name;
                    if (sheet.Name.Contains(ltcCustomSheetName) == true)
                    {
                        //Utlity.Log("Skipping Custom LTC Sheet: " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }
                // 29 - SEPT, Ignore Sheets that Are not changed by the User/External Link
                if (Ribbon1.dialog.syncAllCb.Checked == false) //Change requested by LTC to sync all sheets 25-Dec-2019
                {
                    if (Utlity.ModSheetsInSession != null && Utlity.ModSheetsInSession.Contains(sheet.Name) == false)
                    {
                        //Utlity.Log(sheet.Name + "::::" + " Is not Modified By User, Skipping It", logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }


                List<Variable> variableListForSheet = new List<Variable>();
                
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                try
                {
                    if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        // 10 - SEPT - NOT READING SHEETS WHICH ARE HIDDEN. PERFORMANCE OVER HEAD.
                        partEnablementDictionary.Add(sheet.Name, false);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                    else
                    {
                        partEnablementDictionary.Add(sheet.Name, true);
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }

                Utlity.Log("MODIFIED " + sheet.Name, logFilePath);

                Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                object[,] values = (object[,])xlRange.Value2;
                for (int i = 1; i <= values.GetLength(0); i++)                
                //for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    if (i == 1)
                        continue;
                    try
                    {
                        //Utlity.Log("Iteration: " + i.ToString(), logFilePath);

                        Variable varr = new Variable();
                        //if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                        if (values[i, 1] != null)
                        {
                            try
                            {
                                //varr.name = xlRange.Cells[i, 1].Value2;
                                varr.name = Convert.ToString(values[i, 1]);
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
                        //if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2] != null)
                        if (values[i, 2] != null)
                        {
                            try
                            {
                                //varr.systemName = xlRange.Cells[i, 2].Value2;
                                varr.systemName = Convert.ToString(values[i, 2]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("systemName" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        if (values[i, 3] != null)
                        {
                            try
                            {
                                //varr.value = xlRange.Cells[i, 3].Value2;
                                varr.value = Convert.ToString(values[i, 3]);
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("value" + ex.Message, logFilePath);
                                double value = xlRange.Cells[i, 3].Value2;
                                varr.value = value.ToString("0.######");
                            }
                        }
                        //if (xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4] != null)
                        if (values[i, 4] != null)
                        {
                            try
                            {
                                //varr.unit = xlRange.Cells[i, 4].Value2;
                                varr.unit = Convert.ToString(values[i, 4]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("unit" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                        if (values[i, 5] != null)
                        {
                            try
                            {
                                //varr.rangeLow = xlRange.Cells[i, 5].Value2;
                                varr.rangeLow = Convert.ToString(values[i, 5]);
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("rangeLow" + ex.Message, logFilePath);
                                double value = xlRange.Cells[i, 5].Value2;
                                varr.rangeLow = value.ToString("0.######");

                            }
                        }
                        //if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        if (values[i, 6] != null)
                        {
                            try
                            {
                                //varr.rangeHigh = xlRange.Cells[i, 6].Value2;
                                varr.rangeHigh = Convert.ToString(values[i, 6]);
                            }
                            catch (Exception ex)
                            {
                                double value = xlRange.Cells[i, 6].Value2;
                                varr.rangeHigh = value.ToString("0.######");
                                //Utlity.Log("rangeHigh" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        if (values[i, 7] != null)
                        {
                            try
                            {
                                //int result = 0;
                                //Double rangeCondition = xlRange.Cells[i, 7].Value2;
                                //result = (int)rangeCondition;
                                //varr.rangeCondition = result;
                                varr.rangeCondition = Convert.ToInt32(values[i, 7]);

                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Trouble Parsing rangeCondition: " + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                        if (values[i, 8] != null)
                        {
                            try
                            {
                                //varr.Formula = xlRange.Cells[i, 8].Value2;
                                varr.Formula = Convert.ToString(values[i, 8]);
                            }
                            catch (Exception ex)
                            {
                                //Utlity.Log("Formula: " + ex.Message, logFilePath);
                                int Formula = xlRange.Cells[i, 8].Value2;
                                varr.Formula = Formula.ToString();
                            }
                        }
                        //if (xlRange.Cells[i, 9].Value2 != null && xlRange.Cells[i, 9] != null)
                        if (values[i, 9] != null)
                        {
                            try
                            {
                                //varr.PartName = xlRange.Cells[i, 9].Value2;
                                varr.PartName = Convert.ToString(values[i, 9]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("PartName" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 10].Value2 != null && xlRange.Cells[i, 10] != null)
                        if (values[i, 10] != null)
                        {
                            try
                            {
                                //varr.AddPartToTemplate = xlRange.Cells[i, 10].Value2;
                                varr.AddPartToTemplate = Convert.ToString(values[i, 10]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AddPartToTemplate" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 11] != null)
                        if (values[i, 11] != null)
                        {
                            try
                            {
                                //varr.AddVarToTemplate = xlRange.Cells[i, 11].Value2;
                                varr.AddVarToTemplate = Convert.ToBoolean(values[i, 11]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AddVarToTemplate" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 12].Value2 != null && xlRange.Cells[i, 12] != null)
                        if (values[i, 12] != null)
                        {
                            try
                            {
                                //varr.LOV = xlRange.Cells[i, 12].Value2;
                                varr.LOV = Convert.ToString(values[i, 12]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("LOV" + ex.Message, logFilePath);
                            }
                        }
                        //if (xlRange.Cells[i, 13].Value2 != null && xlRange.Cells[i, 13] != null)
                        if (values[i, 13] != null)
                        {
                            try
                            {
                                //varr.variableType = xlRange.Cells[i, 13].Value2;
                                varr.variableType = Convert.ToString(values[i, 13]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("variableType" + ex.Message, logFilePath);
                            }
                        }
                        if (values[i, 14] != null)
                        {
                            try
                            {
                                
                                //varr.DefaultValue = Convert.ToString(values[i, 14]);
                                varr.DefaultValue = Convert.ToString(values[i, 14]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("value" + ex.Message, logFilePath);
                                
                            }
                        }
                        //if (xlRange.Cells[i, 15].Value2 != null && xlRange.Cells[i, 15] != null)
                        if (values[i, 15] != null)
                        {
                            try
                            {
                                //varr.UnitType = xlRange.Cells[i, 15].Value2;
                                varr.UnitType = Convert.ToString(values[i, 15]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("UnitType" + ex.Message, logFilePath);

                            }
                        }
                        ALLvariablesList.Add(varr);
                        variableListForSheet.Add(varr);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                    }


                }
                Utlity.Log("Releasing Range", logFilePath);
                Marshal.ReleaseComObject(xlRange);
                xlRange = null;



                variableDictionary.Add(sheet.Name, variableListForSheet);

                Utlity.Log("Releasing sheet", logFilePath);
                Marshal.ReleaseComObject(sheet);
                //Utlity.Log(sheet.Name + " is Done", logFilePath);  
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Utlity.Log("Releasing Sheets", logFilePath);
            Marshal.ReleaseComObject(sheets);
            sheets = null;

            
            Utlity.Log("Releasing workbook", logFilePath);
            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;
            Utlity.Log("----------------------------------------------------------", logFilePath);
            
            return ALLvariablesList;

        }

        

       
    }
}
