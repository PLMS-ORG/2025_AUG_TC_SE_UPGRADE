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
    class ExcelData
    {
        public static List<String> occurenceList = new List<string>();
        public static List<Variable> ALLvariablesList = new List<Variable>();
        public static List<BOMLine> bomLineList = new List<BOMLine>();
        public static Dictionary<String, List<Variable>> variableDictionary = new Dictionary<string, List<Variable>>
            ();
        public static Dictionary<String, bool> partEnablementDictionary = new Dictionary<string, bool>();
        //public static Dictionary<String, String> occurencePathDictionary = new Dictionary<string, string>();

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

        //public static Dictionary<String, String> getOcurrencePathDictionary()
        //{
        //    return occurencePathDictionary;
        //}

        public static List<String> getOcurrenceList()
        {
            return occurenceList;
        }
        public static List<BOMLine> getBomLineList()
        {
            return bomLineList;
        }

        public static List<Variable> readOccurenceVariablesFromTemplateExcel(String xlFilePath,String logFilePath)
        {
            ALLvariablesList.Clear();
            partEnablementDictionary.Clear();
            variableDictionary.Clear();
            occurenceList.Clear();

            Utlity.Log("----------------------------------------------------------", logFilePath);
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
                return null;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return null;
            }
            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            //Utlity.Log("WorkSheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
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
               
                 List<Variable> variableListForSheet = new List<Variable>();
                
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true )
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                try
                {
                    Utlity.Log(sheet.Name, logFilePath);
                    if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        partEnablementDictionary.Add(sheet.Name, false);
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
                        ALLvariablesList.Add(varr);
                        variableListForSheet.Add(varr);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message, logFilePath);
                    }

                    
                }
                Marshal.ReleaseComObject(xlRange);
                xlRange = null;
                variableDictionary.Add(sheet.Name, variableListForSheet);

                Marshal.ReleaseComObject(sheet);               

                //Utlity.Log(sheet.Name + " is Done", logFilePath);  
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background            
            //Marshal.ReleaseComObject(xlWorksheet);

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
            return ALLvariablesList;

        }


        public static List<Variable> readOccurenceVariablesFromTemplateExcelFast(String xlFilePath, String logFilePath)
        {
            ALLvariablesList.Clear();
            partEnablementDictionary.Clear();
            variableDictionary.Clear();
            occurenceList.Clear();

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            if (f.Exists == true)
            {
                workbooks = xlApp.Workbooks;
                //Utlity.Log("File Already Exists", logFilePath);
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
                return null;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return null;
            }
            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            //Utlity.Log("WorkSheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
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

                List<Variable> variableListForSheet = new List<Variable>();

                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                try
                {
                    //Utlity.Log(sheet.Name, logFilePath);
                    if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        partEnablementDictionary.Add(sheet.Name, false);
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

                object[,] values = (object[,])xlRange.Value2;                

                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    if (i == 1)
                        continue;
                    try
                    {
                        //Utlity.Log("Iteration: " + i.ToString(), logFilePath);

                        Variable varr = new Variable();
                       // if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                        if (values[i,1] !=null)
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
                                Utlity.Log("value" + ex.Message, logFilePath);
                                //double value = xlRange.Cells[i, 3].Value2;
                                //varr.value = value.ToString("0.######");
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
                                Utlity.Log("rangeLow" + ex.Message, logFilePath);
                                //double value = xlRange.Cells[i, 5].Value2;
                                //varr.rangeLow = value.ToString("0.######");

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
                                //double value = xlRange.Cells[i, 6].Value2;
                                //varr.rangeHigh = value.ToString("0.######");
                                Utlity.Log("rangeHigh" + ex.Message, logFilePath);
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
                                Utlity.Log("Formula: " + ex.Message, logFilePath);
                                // Formula = xlRange.Cells[i, 8].Value2;
                                //varr.Formula = Formula.ToString();
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
                                //varr.AddVarToTemplate = (bool)xlRange.Cells[i, 11].Value2;
                                varr.AddVarToTemplate = Convert.ToBoolean(values[i, 11]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AddVarToTemplate" + ex.Message, logFilePath);
                            }
                        }
                       // if (xlRange.Cells[i, 12].Value2 != null && xlRange.Cells[i, 12] != null)
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
                        //if (xlRange.Cells[i, 14].Value2 != null && xlRange.Cells[i, 14] != null)
                        if (values[i, 14] != null)
                        {
                            try
                            {
                                //varr.DefaultValue = xlRange.Cells[i, 14].Value2;
                                varr.DefaultValue = Convert.ToString(values[i, 14]);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("value" + ex.Message, logFilePath);
                                //double DefaultValue = xlRange.Cells[i, 14].Value2;
                                //varr.DefaultValue = DefaultValue.ToString("0.######");
                            }
                        }

                       // if (xlRange.Cells[i, 15].Value2 != null && xlRange.Cells[i, 15] != null)
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
                Marshal.ReleaseComObject(xlRange);
                xlRange = null;
                variableDictionary.Add(sheet.Name, variableListForSheet);

                Marshal.ReleaseComObject(sheet);

                //Utlity.Log(sheet.Name + " is Done", logFilePath);  
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background            
            //Marshal.ReleaseComObject(xlWorksheet);

            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Close(true);


            Marshal.ReleaseComObject(sheets);
            sheets = null;

            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            Marshal.ReleaseComObject(workbooks);
            workbooks = null;

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            xlApp.DisplayAlerts = true;
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            return ALLvariablesList;

        }

        public static void readOccurencePathFromTemplateExcel(String xlFilePath, String logFilePath)
        {
            //occurencePathDictionary.Clear();
            occurenceList.Clear();
            bomLineList.Clear();

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
           
            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (f.Exists == true)
            {
                //Utlity.Log("File Already Exists", logFilePath);
                xlApp.DisplayAlerts = false;
                //xlWorkbook = Workbooks.Open(xlFilePath);
                try
                {
                    xlWorkbook = Workbooks.Open(xlFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = Workbooks.Open(xlFilePath, CorruptLoad: 1);
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
                return ;
            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return ;
            }
            String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //Utlity.Log(sheet.Name, logFilePath);
                    Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i == 1)
                            continue;
                        BOMLine bl = new BOMLine();

                        if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // AbsolutePath
                                filePath = xlRange.Cells[i, 3].Value2;
                                String fileName = Path.GetFileName(filePath);
                                bl.AbsolutePath = filePath;
                                //occurencePathDictionary.Add(fileName, filePath);
                                String fileWithExtn = Path.GetFileName(fileName);
                                if (fileWithExtn != null && fileWithExtn.Equals("") == false)
                                {
                                    if (occurenceList.Contains(fileWithExtn) == false)
                                    {
                                        occurenceList.Add(fileWithExtn);
                                    }
                                    else
                                    {
                                        Utlity.Log("Skipping, Added to occurenceList Already," + fileWithExtn, logFilePath);                                       
                                        continue;
                                    }
                                }
                                //Utlity.Log(fileName, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AbsolutePath" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.AbsolutePath = "";
                        }

                        if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // FullName
                                filePath = xlRange.Cells[i, 7].Value2;
                                bl.FullName = filePath;
                                //Utlity.Log(filePath, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("FullName" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.FullName = "";
                        }

                        if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2] != null)
                        {
                            try
                            {
                                Double Level = 0;
                                // Level
                                Level = (Double)xlRange.Cells[i, 2].Value2;
                                bl.level = Level.ToString();
                                //Utlity.Log(Level.ToString(), logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Level" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.level = "";
                        }

                        if (xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4] != null)
                        {
                            try
                            {
                                String DocNum = "";
                                // DocNum
                                DocNum = xlRange.Cells[i, 4].Value2;
                                bl.DocNum = DocNum;
                                //Utlity.Log(DocNum, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Level" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.DocNum = "";
                        }
                        if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                        {
                            try
                            {
                                String Revision = "";
                                // Revision
                                Revision = xlRange.Cells[i, 5].Value2;
                                bl.Revision = Revision;
                                //Utlity.Log(Revision, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Revision" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.Revision = "";
                        }
                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {
                            try
                            {
                                String Status = "";
                                // Status
                                Status = xlRange.Cells[i, 6].Value2;
                                bl.Status = Status;
                                //Utlity.Log(Status, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Status" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.Status = "";
                        }

                        bomLineList.Add(bl);

                    }
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;

                    Marshal.ReleaseComObject(sheet);
                }
                else
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }


            }

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            
            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Close(true);
            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;
            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;
            Utlity.Log("----------------------------------------------------------", logFilePath);
            xlApp.DisplayAlerts = true;
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

        }


        public static void readOccurencePathFromTemplateExcelFast(String xlFilePath, String logFilePath)
        {
            //occurencePathDictionary.Clear();
            occurenceList.Clear();
            bomLineList.Clear();

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (f.Exists == true)
            {
                //Utlity.Log("File Already Exists", logFilePath);
                xlApp.DisplayAlerts = false;
                //xlWorkbook = Workbooks.Open(xlFilePath);
                try
                {
                    xlWorkbook = Workbooks.Open(xlFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = Workbooks.Open(xlFilePath, CorruptLoad: 1);
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
            String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {

                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //Utlity.Log(sheet.Name, logFilePath);
                    Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                    //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);

                    object[,] values = (object[,])xlRange.Value2;
                    for (int i = 1; i <= values.GetLength(0); i++)
                    //for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i == 1)
                            continue;
                        BOMLine bl = new BOMLine();

                        //if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        if (values[i, 3] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // AbsolutePath
                                //filePath = xlRange.Cells[i, 3].Value2;
                                filePath = Convert.ToString(values[i, 3]);
                                String fileName = Path.GetFileName(filePath);
                                bl.AbsolutePath = filePath;
                                //occurencePathDictionary.Add(fileName, filePath);
                                String fileWithExtn = Path.GetFileName(fileName);
                                if (fileWithExtn != null && fileWithExtn.Equals("") == false)
                                {
                                    if (occurenceList.Contains(fileWithExtn) == false)
                                    {
                                        occurenceList.Add(fileWithExtn);
                                    }
                                    else
                                    {
                                        Utlity.Log("Skipping, Added to occurenceList Already," + fileWithExtn, logFilePath);
                                        continue;
                                    }
                                }
                                //Utlity.Log(fileName, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AbsolutePath" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.AbsolutePath = "";
                        }

                        //if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        if (values[i, 7] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // FullName
                                //filePath = xlRange.Cells[i, 7].Value2;
                                filePath = Convert.ToString(values[i, 7]);
                                bl.FullName = filePath;
                                //Utlity.Log(filePath, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("FullName" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.FullName = "";
                        }

                        //if (xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2] != null)
                        if (values[i, 2] != null)
                        {
                            try
                            {
                                bl.level = Convert.ToString(values[i, 2]);
                                //Double Level = 0;
                                //// Level
                                //Level = (Double)xlRange.Cells[i, 2].Value2;
                                //bl.level = Level.ToString();
                                //Utlity.Log(Level.ToString(), logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Level" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.level = "";
                        }

                        //if (xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4] != null)
                        if (values[i, 4] != null)
                        {
                            try
                            {
                                bl.DocNum = Convert.ToString(values[i, 4]);
                                //String DocNum = "";
                                //// DocNum
                                //DocNum = xlRange.Cells[i, 4].Value2;
                                //bl.DocNum = DocNum;
                                //Utlity.Log(DocNum, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Level" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.DocNum = "";
                        }

                        //if (xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5] != null)
                        if (values[i, 5] != null)
                        {
                            try
                            {
                                bl.Revision = Convert.ToString(values[i, 5]);
                                //String Revision = "";
                                //// Revision
                                //Revision = xlRange.Cells[i, 5].Value2;
                                //bl.Revision = Revision;
                                //Utlity.Log(Revision, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Revision" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.Revision = "";
                        }

                        //if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        if (values[i, 6] != null)
                        {
                            try
                            {
                                bl.Status = Convert.ToString(values[i, 6]);
                                //String Status = "";
                                //// Status
                                //Status = xlRange.Cells[i, 6].Value2;
                                //bl.Status = Status;
                                //Utlity.Log(Status, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Status" + ex.Message, logFilePath);
                            }
                        }
                        else
                        {
                            bl.Status = "";
                        }

                        bomLineList.Add(bl);

                    }
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;

                    Marshal.ReleaseComObject(sheet);
                }
                else
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }


            }

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Close(true);
            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;
            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;
            Utlity.Log("----------------------------------------------------------", logFilePath);
            xlApp.DisplayAlerts = true;
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

        }

       
    }
}
