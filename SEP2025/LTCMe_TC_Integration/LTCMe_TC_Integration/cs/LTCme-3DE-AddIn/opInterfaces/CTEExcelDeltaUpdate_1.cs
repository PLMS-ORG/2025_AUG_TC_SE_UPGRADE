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
    class CTEExcelDeltaUpdate_1
    {
        /** Step-1, Open the New XL and Copy the Input_Data sheet from Old to the New One.
         *  Step-2, Iterate the Sheets,
         *  Step-3, Copy the Formula Alone from Old Part Sheet to the New Part Sheet
         *  Step-4, Copy the Formula Aone from Old Component Sheet to the New Component Sheet
         **/

        public static void ProcessNewTemplate(String NewExcelTemplateFilePath, String OldExcelTemplateFilePath, String logFilePath)
        {

            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            CopySheetNew(NewExcelTemplateFilePath, OldExcelTemplateFilePath, ltcCustomSheetName, logFilePath);
            //Utlity.Log("ProcessNewTemplate: CopyFormula...Started", logFilePath);
            CopyFormulaCorrect(NewExcelTemplateFilePath, OldExcelTemplateFilePath, logFilePath);
            //Copying images from old excel to new excel
            Utility.Log("Copying images from old excel " + OldExcelTemplateFilePath + " to new excel " + NewExcelTemplateFilePath, logFilePath);
            CopyImageFromOneExceltoAnother(OldExcelTemplateFilePath, NewExcelTemplateFilePath, logFilePath);

        }

        // copies the formula from old sheet to new sheet
        // copies formula from features Tab
        // copies formula from part tab
        private static void CopyFormula(string NewExcelTemplateFilePath, string OldExcelTemplateFilePath, string logFilePath)
        {
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null) return;

            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = XlWindowState.xlNormal;

            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (Workbooks == null) return;

            FileInfo f = new FileInfo(OldExcelTemplateFilePath);
            if (f == null) return;
            //-------------Old XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook OldxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + OldExcelTemplateFilePath, logFilePath);
                //OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                try
                {
                    OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        OldxlWorkbook = Workbooks.Open(OldExcelTemplateFilePath, CorruptLoad: 1);
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
                return;
            }
            if (OldxlWorkbook == null)
            {
                Utlity.Log("OldxlWorkbook is NULL", logFilePath);
                return;
            }

            //FileInfo f1 = new FileInfo(NewExcelTemplateFilePath);
            //if (f1 == null) return;
            ////-------------New XL WorkBook--------------------------
            //Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook = null;
            //if (f1.Exists == true)
            //{
            //    //Utlity.Log("Opening..," + NewExcelTemplateFilePath, logFilePath);
            //    NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
            //}
            //else
            //{
            //    Utlity.Log("File Does Not Exist,", logFilePath);
            //    return;
            //}
            //if (NewxlWorkbook == null)
            //{
            //    Utlity.Log("NewxlWorkbook is NULL", logFilePath);
            //    return;
            //}

            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            Microsoft.Office.Interop.Excel.Sheets Oldsheets = OldxlWorkbook.Worksheets;
            if (Oldsheets == null) return;
            foreach (Microsoft.Office.Interop.Excel.Worksheet OldSheet in OldxlWorkbook.Worksheets)
            {
                //Utlity.Log(OldSheet.Name, logFilePath);
                if (OldSheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(OldSheet);
                    continue;
                }

                if (OldSheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                {
                    Marshal.ReleaseComObject(OldSheet);
                    continue;
                }

                if (OldSheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    // Features ----
                    UpdateFormulaInNewTemplate(NewExcelTemplateFilePath, logFilePath, OldSheet);


                }
                else
                {
                    UpdateFormulaInNewTemplate(NewExcelTemplateFilePath, logFilePath, OldSheet);

                }


                Marshal.ReleaseComObject(OldSheet);
            }

            // Old Workbook - Release
            OldxlWorkbook.Save();
            OldxlWorkbook.Close(true);

            Marshal.ReleaseComObject(Oldsheets);
            Oldsheets = null;

            Marshal.ReleaseComObject(OldxlWorkbook);
            OldxlWorkbook = null;

            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

        }


        // copies the formula from old sheet to new sheet
        // copies formula from features Tab
        // copies formula from part tab
        private static void CopyFormulaCorrect(string NewExcelTemplateFilePath, string OldExcelTemplateFilePath, string logFilePath)
        {
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null) return;

            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = XlWindowState.xlNormal;
            //xlApp.Calculation = XlCalculation.xlCalculationManual;

            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (Workbooks == null) return;

            FileInfo f = new FileInfo(OldExcelTemplateFilePath);
            if (f == null) return;
            //-------------Old XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook OldxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + OldExcelTemplateFilePath, logFilePath);
                //OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                try
                {
                    OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath, CorruptLoad: 1);
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
                return;
            }
            if (OldxlWorkbook == null)
            {
                Utlity.Log("OldxlWorkbook is NULL", logFilePath);
                return;
            }

            FileInfo f1 = new FileInfo(NewExcelTemplateFilePath);
            if (f1 == null) return;
            //-------------New XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook = null;
            if (f1.Exists == true)
            {
                //Utlity.Log("Opening..," + NewExcelTemplateFilePath, logFilePath);
                //NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                try
                {
                    NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Utlity.Log("File Does Not Exist,", logFilePath);
                return;
            }
            if (NewxlWorkbook == null)
            {
                Utlity.Log("NewxlWorkbook is NULL", logFilePath);
                return;
            }

            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            Microsoft.Office.Interop.Excel.Sheets Oldsheets = OldxlWorkbook.Worksheets;
            if (Oldsheets == null) return;
            foreach (Microsoft.Office.Interop.Excel.Worksheet OldSheet in OldxlWorkbook.Worksheets)
            {
                //Utlity.Log(OldSheet.Name, logFilePath);
                if (OldSheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    UpdateMasterAssemblyFormulaInNewTemplateCorrect(xlApp, NewxlWorkbook, logFilePath, OldSheet);

                    Marshal.ReleaseComObject(OldSheet);
                    continue;
                }

                if (OldSheet.Name.Contains(ltcCustomSheetName) == true)
                {
                    Marshal.ReleaseComObject(OldSheet);
                    continue;
                }

                if (OldSheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                {
                    Marshal.ReleaseComObject(OldSheet);
                    continue;
                }

                if (OldSheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    // Features ----
                    //UpdateFormulaInNewTemplate(NewExcelTemplateFilePath, logFilePath, OldSheet);
                    UpdateFeatureFormulaInNewTemplateCorrect(xlApp, NewxlWorkbook, logFilePath, OldSheet);

                }
                else
                {
                    //UpdateFormulaInNewTemplate(NewExcelTemplateFilePath, logFilePath, OldSheet);
                    UpdateFormulaInNewTemplateCorrect(xlApp, NewxlWorkbook, logFilePath, OldSheet);

                }


                Marshal.ReleaseComObject(OldSheet);
            }

            // Old Workbook - Release
            OldxlWorkbook.Save();
            OldxlWorkbook.Close(true);

            Marshal.ReleaseComObject(Oldsheets);
            Oldsheets = null;

            Marshal.ReleaseComObject(OldxlWorkbook);
            OldxlWorkbook = null;

            NewxlWorkbook.Save();
            NewxlWorkbook.Close(true);

            Marshal.ReleaseComObject(NewxlWorkbook);
            NewxlWorkbook = null;

            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;

            //xlApp.Calculation = XlCalculation.xlCalculationAutomatic;
            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

        }

        private static void UpdateMasterAssemblyFormulaInNewTemplateCorrect(_Application xlApp, Workbook NewxlWorkbook, string logFilePath, Worksheet OldSheet)
        {
            Utlity.Log("UpdateMasterAssemblyFormulaInNewTemplateCorrect: ", logFilePath);
            //Utlity.Log("NewExcelTemplateFilePath: " + NewExcelTemplateFilePath, logFilePath);

            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            try
            {
                sheet = NewxlWorkbook.Sheets["MASTER ASSEMBLY"];
            }
            catch (Exception ex)
            {
                Utlity.Log("could not find sheet in New Template : " + OldSheet.Name, logFilePath);
                Utlity.Log("could not find sheet in New Template : " + ex.Message, logFilePath);
                return;
            }
            if (sheet == null)
            {
                Utlity.Log("could not find sheet in New Template : " + OldSheet.Name, logFilePath);
                return;
            }

            try
            {
                sheet.Select();
            }
            catch (Exception ex)
            {
                Utlity.Log("UpdateMasterAssemblyFormulaInNewTemplateCorrect: Sheet Select Failed due to Exception : " + ex.Message, logFilePath);
            }


            Dictionary<String, String> MasterAssemblyFormulaDictionary = ReadFormulaInMasterAssemblyTabCorrect(logFilePath, OldSheet);

            if (MasterAssemblyFormulaDictionary == null || MasterAssemblyFormulaDictionary.Count == 0)
            {
                Utlity.Log("MasterAssemblyFormulaDictionary is Empty : " + sheet.Name, logFilePath);
                Marshal.ReleaseComObject(sheet);
                return;
            }

            UpdateFormulaInMasterAssemblySheetCorrect(xlApp, NewxlWorkbook, sheet, MasterAssemblyFormulaDictionary, logFilePath);
            MasterAssemblyFormulaDictionary.Clear();

            NewxlWorkbook.Save();
            //NewWorkbook.Close(true);
            Marshal.ReleaseComObject(sheet);
            sheet = null;
        }

        private static void UpdateFormulaInMasterAssemblySheetCorrect(_Application xlApp, Workbook NewxlWorkbook, Worksheet sheet, Dictionary<string, string> MasterSheetFormulaDictionary, string logFilePath)
        {
            if (sheet == null) return;
            if (MasterSheetFormulaDictionary == null || MasterSheetFormulaDictionary.Count == 0) return;

            Utlity.Log("UpdateFormulaInMasterAssemblySheetCorrect: ", logFilePath);
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            object[,] values = (object[,])xlRange.Value2;



            if (xlRange == null) return;
            xlApp.DisplayFormulaBar = true;




            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (i == 1)
                    continue;

                String PartName = "";
                if (values[i, 1] != null)
                {
                    PartName = Convert.ToString(values[i, 1]);

                }


                if (PartName != null && PartName.Equals("") == false)
                {
                    String Formula = "";
                    MasterSheetFormulaDictionary.TryGetValue(PartName, out Formula);

                    if (Formula != null && Formula.Equals("") == false)
                    {
                        values[i, 6] = Formula;

                    }
                }

            }

            int rows = values.GetLength(0);
            int columns = values.GetLength(1);

            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.get_Range("A1", Type.Missing);
            range = range.get_Resize(rows, columns);
            // Assign the Array to the Range in one shot:
            range.set_Value(Type.Missing, values);
            Marshal.ReleaseComObject(range);
            range = null;

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;
        }

        private static Dictionary<string, string> ReadFormulaInMasterAssemblyTabCorrect(string logFilePath, Worksheet OldxlWorkSheet)
        {
            Dictionary<String, String> MasterAssemblyFormulaDictionary = new Dictionary<string, string>();
            Utlity.Log("ReadFormulaInMasterAssemblyTabCorrect: ", logFilePath);

            try
            {
                if (OldxlWorkSheet == null) return null;
            }
            catch (Exception ex)
            {
                Utlity.Log("OldxlWorkSheet is NULL..." + ex.Message, logFilePath);
                return null;
            }

            Microsoft.Office.Interop.Excel.Range xlRange = OldxlWorkSheet.UsedRange;

            object[,] values = (object[,])xlRange.Formula;
            object[,] values1 = (object[,])xlRange.Value2;

            if (xlRange == null) return null;

            if (values == null)
            {
                Utlity.Log("ReadFormulaInMasterAssemblyTabCorrect... No Formula in the Old Work sheet...", logFilePath);
                return null;
            }

            if (values1 == null)
            {
                Utlity.Log("ReadFormulaInMasterAssemblyTabCorrect... No Values in the Old Work sheet...", logFilePath);
                return null;
            }


            for (int i = 1; i <= values1.GetLength(0); i++)
            {
                if (i == 1)
                    continue;
                try
                {

                    String FullName = "";
                    String StatusFormulaValue = "";


                    if (values1[i, 1] != null)
                    {
                        try
                        {
                            FullName = Convert.ToString(values1[i, 1]);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("PartName: " + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        Utlity.Log("FullName not Found.. :" + i, logFilePath);
                        continue;
                    }


                    if (values[i, 6] != null)
                    {
                        try
                        {
                            StatusFormulaValue = Convert.ToString(values[i, 6]);



                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("StatusFormulaValue: " + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        Utlity.Log("StatusFormulaValue : " + i, logFilePath);
                        continue;
                    }



                    if (StatusFormulaValue != null && StatusFormulaValue.Equals("") == false)
                    {
                        if (StatusFormulaValue.StartsWith("=") == true)
                        {

                            if (FullName != null && FullName.Equals("") == false && StatusFormulaValue != null && StatusFormulaValue.Equals("") == false)
                            {

                                MasterAssemblyFormulaDictionary.Add(FullName, StatusFormulaValue);
                            }
                        }

                    }


                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }


            }
            Marshal.ReleaseComObject(xlRange);
            xlRange = null;


            return MasterAssemblyFormulaDictionary;
        }





        public static void CopySheet(String NewExcelTemplateFilePath, String OldExcelTemplateFilePath, String
            SheetNameToCopy, String logFilePath)
        {
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null) return;

            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = XlWindowState.xlNormal;

            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (Workbooks == null) return;
            FileInfo f = new FileInfo(NewExcelTemplateFilePath);
            if (f == null) return;
            //-------------New XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + NewExcelTemplateFilePath, logFilePath);
                //NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                try
                {
                    NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath, CorruptLoad: 1);
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
                return;
            }
            if (NewxlWorkbook == null)
            {
                Utlity.Log("NewxlWorkbook is NULL", logFilePath);
                return;
            }

            //-------------Old XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook OldxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + OldExcelTemplateFilePath, logFilePath);
                //OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                try
                {
                    OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath, CorruptLoad: 1);
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
                return;
            }
            if (OldxlWorkbook == null)
            {
                Utlity.Log("OldxlWorkbook is NULL", logFilePath);
                return;
            }


            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);

            try
            {
                if (OldxlWorkbook.Worksheets[ltcCustomSheetName] != null)
                {
                    Microsoft.Office.Interop.Excel.Sheets Oldsheets = OldxlWorkbook.Worksheets;
                    if (Oldsheets != null)
                    {
                        Microsoft.Office.Interop.Excel._Worksheet _OldworkSheet = OldxlWorkbook.Worksheets[ltcCustomSheetName];
                        Microsoft.Office.Interop.Excel._Worksheet _newWorkSheet = NewxlWorkbook.Worksheets[NewxlWorkbook.Worksheets.Count];
                        if (_OldworkSheet != null)
                        {
                            if (_newWorkSheet != null)
                            {
                                Utlity.Log("Copy...", logFilePath);
                                _OldworkSheet.Activate();
                                _OldworkSheet.Select();
                                _OldworkSheet.Copy(Type.Missing, _newWorkSheet);

                            }
                        }

                        Marshal.ReleaseComObject(_OldworkSheet);
                        _OldworkSheet = null;

                        Marshal.ReleaseComObject(_newWorkSheet);
                        _newWorkSheet = null;

                        Marshal.ReleaseComObject(Oldsheets);
                        Oldsheets = null;

                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("Could Not Copy Input-Data Sheet from Old Template to New Template..", logFilePath);
                //Utlity.Log(ex.Message, logFilePath);
            }


            OldxlWorkbook.Save();
            OldxlWorkbook.Close(true);
            NewxlWorkbook.Save();
            NewxlWorkbook.SaveAs(NewExcelTemplateFilePath, AccessMode: XlSaveAsAccessMode.xlShared);
            NewxlWorkbook.Close(true);

            Marshal.ReleaseComObject(NewxlWorkbook);
            NewxlWorkbook = null;

            Marshal.ReleaseComObject(OldxlWorkbook);
            OldxlWorkbook = null;

            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
        }

        public static void CopySheetNew(String NewExcelTemplateFilePath, String OldExcelTemplateFilePath, String
            SheetNameToCopy, String logFilePath)
        {
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null) return;

            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = XlWindowState.xlNormal;

            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (Workbooks == null) return;
            FileInfo f = new FileInfo(NewExcelTemplateFilePath);
            if (f == null) return;
            //-------------New XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + NewExcelTemplateFilePath, logFilePath);
                //NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                try
                {
                    NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath, CorruptLoad: 1);
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
                return;
            }
            if (NewxlWorkbook == null)
            {
                Utlity.Log("NewxlWorkbook is NULL", logFilePath);
                return;
            }

            //-------------Old XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook OldxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + OldExcelTemplateFilePath, logFilePath);
                //OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                try
                {
                    OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        OldxlWorkbook = xlApp.Workbooks.Open(OldExcelTemplateFilePath, CorruptLoad: 1);
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
                return;
            }
            if (OldxlWorkbook == null)
            {
                Utlity.Log("OldxlWorkbook is NULL", logFilePath);
                return;
            }


            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);

            try
            {
                if (OldxlWorkbook.Worksheets[ltcCustomSheetName] != null)
                {
                    Microsoft.Office.Interop.Excel.Sheets Oldsheets = OldxlWorkbook.Worksheets;
                    if (Oldsheets != null)
                    {
                        Microsoft.Office.Interop.Excel._Worksheet _newWorkSheet = NewxlWorkbook.Worksheets[NewxlWorkbook.Worksheets.Count];
                        foreach (Microsoft.Office.Interop.Excel._Worksheet OldSheet in OldxlWorkbook.Worksheets)
                        {
                            if (OldSheet.Name.Contains(ltcCustomSheetName) == true)
                            {
                                if (OldSheet != null)
                                {
                                    if (_newWorkSheet != null)
                                    {
                                        Utlity.Log("Copy..." + OldSheet.Name, logFilePath);
                                        OldSheet.Activate();
                                        OldSheet.Select();
                                        OldSheet.Copy(Type.Missing, _newWorkSheet);

                                    }
                                }
                            }
                            Marshal.ReleaseComObject(OldSheet);
                        }




                        Marshal.ReleaseComObject(_newWorkSheet);
                        _newWorkSheet = null;

                        Marshal.ReleaseComObject(Oldsheets);
                        Oldsheets = null;

                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("Could Not Copy Input-Data Sheet from Old Template to New Template..", logFilePath);
                //Utlity.Log(ex.Message, logFilePath);
            }


            OldxlWorkbook.Save();
            OldxlWorkbook.Close(true);
            NewxlWorkbook.Save();
            NewxlWorkbook.SaveAs(NewExcelTemplateFilePath, AccessMode: XlSaveAsAccessMode.xlShared);
            NewxlWorkbook.Close(true);

            Marshal.ReleaseComObject(NewxlWorkbook);
            NewxlWorkbook = null;

            Marshal.ReleaseComObject(OldxlWorkbook);
            OldxlWorkbook = null;

            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
        }

        // Old sheet Processing - Reading the Old sheet and Creating a Dictionary`
        public static Dictionary<String, String> ReadInputDataFormula(String PartSheet, String logFilePath,
            Microsoft.Office.Interop.Excel._Worksheet OldxlWorkSheet)
        {
            Dictionary<String, String> VariableFormulaDictionary = new Dictionary<string, string>();
            //Utlity.Log("ReadInputDataFormula: ", logFilePath);

            try
            {
                if (OldxlWorkSheet == null) return null;
            }
            catch (Exception ex)
            {
                Utlity.Log("OldxlWorkSheet is NULL..." + ex.Message, logFilePath);
                return null;
            }

            Microsoft.Office.Interop.Excel.Range xlRange = OldxlWorkSheet.UsedRange;
            if (xlRange == null) return null;

            //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;
                try
                {
                    //Utlity.Log("Iteration: " + i.ToString(), logFilePath);
                    String VariableName = "";
                    String VariableFormulaValue = "";

                    if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                    {
                        try
                        {

                            VariableName = xlRange.Cells[i, 1].Value2;
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);

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

                    if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                    {
                        try
                        {
                            VariableFormulaValue = getCellFormula(xlRange, i, 3);
                            //VariableFormulaValue = getCellValue(xlRange, i, 3);
                            //VariableFormulaValue = xlRange.Cells[i, 3].Value2;
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("getCellFormula: " + ex.Message, logFilePath);
                            //double value = xlRange.Cells[i, 3].Value2;
                            //VariableFormulaValue = value.ToString("0.######");
                        }
                    }

                    if (VariableFormulaValue != null && VariableFormulaValue.Equals("") == false)
                    {
                        if (VariableFormulaValue.StartsWith("=") == true)
                        {
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                            VariableFormulaDictionary.Add(VariableName, VariableFormulaValue);
                        }

                    }


                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }


            }
            Marshal.ReleaseComObject(xlRange);
            xlRange = null;


            return VariableFormulaDictionary;

        }


        public static Dictionary<String, String> ReadInputDataFormulaCorrect(String PartSheet, String logFilePath,
            Microsoft.Office.Interop.Excel._Worksheet OldxlWorkSheet)
        {
            Dictionary<String, String> VariableFormulaDictionary = new Dictionary<string, string>();
            Utlity.Log("ReadInputDataFormula: " + PartSheet, logFilePath);

            try
            {
                if (OldxlWorkSheet == null) return null;
            }
            catch (Exception ex)
            {
                Utlity.Log("OldxlWorkSheet is NULL..." + ex.Message, logFilePath);
                return null;
            }



            Microsoft.Office.Interop.Excel.Range xlRange = OldxlWorkSheet.UsedRange;
            object[,] values = (object[,])xlRange.Formula;

            object[,] values1 = (object[,])xlRange.Value2;

            if (xlRange == null) return null;

            if (values == null)
            {
                Utlity.Log("ReadInputDataFormulaCorrect... No Formula in the Old Work sheet...", logFilePath);
                return null;
            }

            if (values1 == null)
            {
                Utlity.Log("ReadInputDataFormulaCorrect... No Values in the Old Work sheet...", logFilePath);
                return null;
            }

            for (int i = 1; i <= values1.GetLength(0); i++) // Iterate on the Values, then Take the Formula
            {
                if (i == 1)
                    continue;
                try
                {
                    //Utlity.Log("Iteration: " + i.ToString(), logFilePath);
                    String VariableName = "";
                    String VariableFormulaValue = "";

                    if (values1[i, 1] != null)
                    {
                        try
                        {

                            VariableName = Convert.ToString(values1[i, 1]);
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("name" + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        Utlity.Log("VariableName is null : " + i, logFilePath);
                        continue;
                    }

                    //if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                    if (values[i, 3] != null) // Going to Formula Array here
                    {
                        try
                        {
                            VariableFormulaValue = Convert.ToString(values[i, 3]);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("getCellFormula: " + ex.Message, logFilePath);

                        }
                    }

                    if (VariableFormulaValue != null && VariableFormulaValue.Equals("") == false)
                    {
                        if (VariableFormulaValue.StartsWith("=") == true)
                        {
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                            VariableFormulaDictionary.Add(VariableName, VariableFormulaValue);
                        }

                    }


                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }


            }
            Marshal.ReleaseComObject(xlRange);
            xlRange = null;


            return VariableFormulaDictionary;

        }



        public static string getCellValue(Microsoft.Office.Interop.Excel.Range xlRange, int Row, int Column)
        {
            object cellValue = xlRange.Cells[Row, Column].Value;
            if (cellValue != null)
            {
                return Convert.ToString(cellValue);
            }
            else
            {
                return string.Empty;
            }
        }

        public static String getCellFormula(Microsoft.Office.Interop.Excel.Range xlRange, int Row, int Column)
        {
            Range range = (Range)xlRange.Cells[Row, Column];
            if (range != null)
            {
                return range.Formula;
            }
            return String.Empty;
        }


        // New Sheet Processing - Iterating the sheet and compare the Name with Old Sheet
        // Save the New Sheet.
        public static void UpdateFormulaInNewTemplate(String NewExcelTemplateFilePath, String logFilePath,
            Microsoft.Office.Interop.Excel._Worksheet OldxlWorkSheet)
        {
            //Utlity.Log("UpdateFormulaInNewTemplate: ", logFilePath);
            //Utlity.Log("NewExcelTemplateFilePath: " + NewExcelTemplateFilePath, logFilePath);

            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null) return;
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = XlWindowState.xlNormal;

            Microsoft.Office.Interop.Excel.Workbooks Workbooks = xlApp.Workbooks;
            if (Workbooks == null) return;

            FileInfo f = new FileInfo(NewExcelTemplateFilePath);
            if (f == null) return;

            //-------------New XL WorkBook--------------------------
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook = null;
            if (f.Exists == true)
            {
                //Utlity.Log("Opening..," + NewExcelTemplateFilePath, logFilePath);
                //NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                try
                {
                    NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        NewxlWorkbook = xlApp.Workbooks.Open(NewExcelTemplateFilePath, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Utlity.Log("File Does Not Exist,", logFilePath);
                return;
            }
            if (NewxlWorkbook == null)
            {
                Utlity.Log("NewxlWorkbook is NULL", logFilePath);
                return;
            }

            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
            Microsoft.Office.Interop.Excel.Sheets sheets = NewxlWorkbook.Worksheets;
            if (sheets == null) return;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in NewxlWorkbook.Worksheets)
            {
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        //Utlity.Log("Skipping Custom LTC Sheet: " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }

                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //Utlity.Log("Skipping " + sheet.Name, logFilePath);
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }
                if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                {
                    // 10 - SEPT - NOT READING SHEETS WHICH ARE HIDDEN. PERFORMANCE OVER HEAD.
                    //Utlity.Log("Skipping Hidden Sheet" + sheet.Name, logFilePath);
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Dictionary<String, String> FeatureFormulaDictionary = ReadFormulaInFeaturesTab(OldxlWorkSheet.Name, logFilePath, OldxlWorkSheet);

                    if (FeatureFormulaDictionary == null || FeatureFormulaDictionary.Count == 0)
                    {
                        Utlity.Log("FeatureFormulaDictionary is Empty : " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }

                    UpdateFormulaInFeatureSheetCorrect(xlApp, NewxlWorkbook, sheet, FeatureFormulaDictionary, logFilePath);
                    FeatureFormulaDictionary.Clear();

                }

                //Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals(OldxlWorkSheet.Name) == true)
                {
                    Dictionary<String, String> VariableNameFormulaDictionary = ReadInputDataFormula(OldxlWorkSheet.Name, logFilePath, OldxlWorkSheet);
                    if (VariableNameFormulaDictionary == null || VariableNameFormulaDictionary.Count == 0)
                    {
                        Utlity.Log("VariableNameFormulaDictionary is Empty : " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }

                    UpdateFormulaInNewSheet(xlApp, NewxlWorkbook, sheet, VariableNameFormulaDictionary, logFilePath);
                    VariableNameFormulaDictionary.Clear();
                }
                Marshal.ReleaseComObject(sheet);
            }


            Marshal.ReleaseComObject(sheets);
            sheets = null;

            NewxlWorkbook.Save();
            NewxlWorkbook.Close(true);

            Marshal.ReleaseComObject(NewxlWorkbook);
            NewxlWorkbook = null;

            Marshal.ReleaseComObject(Workbooks);
            Workbooks = null;

            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

        }


        // New Sheet Processing - Iterating the sheet and compare the Name with Old Sheet
        // Save the New Sheet.
        public static void UpdateFormulaInNewTemplateCorrect(Microsoft.Office.Interop.Excel._Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook NewWorkbook, String logFilePath,
            Microsoft.Office.Interop.Excel._Worksheet OldxlWorkSheet)
        {
            Utlity.Log("UpdateFormulaInNewTemplateCorrect: " + OldxlWorkSheet.Name, logFilePath);
            //Utlity.Log("NewExcelTemplateFilePath: " + NewExcelTemplateFilePath, logFilePath);

            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            try
            {
                sheet = NewWorkbook.Sheets[OldxlWorkSheet.Name];
            }
            catch (Exception ex)
            {
                Utlity.Log("could not find sheet in New Template : " + OldxlWorkSheet.Name, logFilePath);
                Utlity.Log("could not find sheet in New Template : " + ex.Message, logFilePath);
                return;
            }
            if (sheet == null)
            {
                Utlity.Log("could not find sheet in New Template : " + OldxlWorkSheet.Name, logFilePath);
                return;
            }

            Boolean SwitchVisibilityFlag = false;
            if (sheet.Visible == XlSheetVisibility.xlSheetHidden || sheet.Visible == XlSheetVisibility.xlSheetVeryHidden)
            {
                SwitchVisibilityFlag = true;
                sheet.Visible = XlSheetVisibility.xlSheetVisible;
            }

            try
            {
                sheet.Select();
            }
            catch (Exception ex)
            {
                Utlity.Log("UpdateFormulaInNewTemplateCorrect, Sheet Select Failed Exception Is : " + ex.Message, logFilePath);
            }

            Dictionary<String, String> VariableNameFormulaDictionary = ReadInputDataFormulaCorrect(OldxlWorkSheet.Name, logFilePath, OldxlWorkSheet);
            if (VariableNameFormulaDictionary == null || VariableNameFormulaDictionary.Count == 0)
            {
                //Utlity.Log("VariableNameFormulaDictionary is Empty : " + sheet.Name, logFilePath);
                if (SwitchVisibilityFlag == true)
                {
                    sheet.Visible = XlSheetVisibility.xlSheetHidden;

                }
                Marshal.ReleaseComObject(sheet);
                return;
            }
            UpdateFormulaInNewSheetCorrect(xlApp, NewWorkbook, sheet, VariableNameFormulaDictionary, logFilePath);
            VariableNameFormulaDictionary.Clear();

            //NewWorkbook.Save();
            //NewWorkbook.Close(true);
            if (SwitchVisibilityFlag == true)
            {
                sheet.Visible = XlSheetVisibility.xlSheetHidden;

            }
            Marshal.ReleaseComObject(sheet);
            sheet = null;

        }

        // New Sheet Processing - Iterating the sheet and compare the Name with Old Sheet
        // Save the New Sheet.
        public static void UpdateFeatureFormulaInNewTemplateCorrect(Microsoft.Office.Interop.Excel._Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook NewWorkbook, String logFilePath,
            Microsoft.Office.Interop.Excel._Worksheet OldxlWorkSheet)
        {
            Utlity.Log("UpdateFeatureFormulaInNewTemplateCorrect: ", logFilePath);
            //Utlity.Log("NewExcelTemplateFilePath: " + NewExcelTemplateFilePath, logFilePath);

            Microsoft.Office.Interop.Excel.Worksheet sheet = null;
            try
            {
                sheet = NewWorkbook.Sheets["FEATURES"];
            }
            catch (Exception ex)
            {
                Utlity.Log("could not find sheet in New Template : " + OldxlWorkSheet.Name, logFilePath);
                Utlity.Log("could not find sheet in New Template : " + ex.Message, logFilePath);
                return;
            }
            if (sheet == null)
            {
                Utlity.Log("could not find sheet in New Template : " + OldxlWorkSheet.Name, logFilePath);
                return;
            }

            try
            {
                sheet.Select();
            }
            catch (Exception ex)
            {
                Utlity.Log("UpdateFeatureFormulaInNewTemplateCorrect: Sheet Select Failed due to Exception : " + ex.Message, logFilePath);
            }


            Dictionary<String, String> FeatureFormulaDictionary = ReadFormulaInFeaturesTabCorrect(OldxlWorkSheet.Name, logFilePath, OldxlWorkSheet);

            if (FeatureFormulaDictionary == null || FeatureFormulaDictionary.Count == 0)
            {
                //Utlity.Log("FeatureFormulaDictionary is Empty : " + sheet.Name, logFilePath);
                Marshal.ReleaseComObject(sheet);
                return;
            }

            UpdateFormulaInFeatureSheetCorrect(xlApp, NewWorkbook, sheet, FeatureFormulaDictionary, logFilePath);
            FeatureFormulaDictionary.Clear();

            NewWorkbook.Save();
            //NewWorkbook.Close(true);
            Marshal.ReleaseComObject(sheet);
            sheet = null;

        }

        private static Dictionary<String, String> ReadFormulaInFeaturesTab(string SheetName, string logFilePath, _Worksheet OldxlWorkSheet)
        {
            Dictionary<String, String> FeatureFormulaDictionary = new Dictionary<string, string>();
            //Utlity.Log("ReadInputDataFormula: ", logFilePath);

            try
            {
                if (OldxlWorkSheet == null) return null;
            }
            catch (Exception ex)
            {
                Utlity.Log("OldxlWorkSheet is NULL..." + ex.Message, logFilePath);
                return null;
            }

            Microsoft.Office.Interop.Excel.Range xlRange = OldxlWorkSheet.UsedRange;
            if (xlRange == null) return null;

            //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;
                try
                {
                    //Utlity.Log("Iteration: " + i.ToString(), logFilePath);
                    String PartName = "";
                    String FeatureSystemName = "";
                    String FeatureFormulaValue = "";

                    if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                    {
                        try
                        {

                            PartName = xlRange.Cells[i, 1].Value2;
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("PartName: " + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        Utlity.Log("Skip Row Number : " + i, logFilePath);
                        continue;
                    }

                    if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                    {
                        try
                        {

                            FeatureSystemName = xlRange.Cells[i, 3].Value2;
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("FeatureSystemName: " + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        Utlity.Log("Skip Row Number : " + i, logFilePath);
                        continue;
                    }

                    if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                    {
                        try
                        {
                            FeatureFormulaValue = getCellFormula(xlRange, i, 8);
                            //VariableFormulaValue = getCellValue(xlRange, i, 3);
                            //VariableFormulaValue = xlRange.Cells[i, 3].Value2;
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("getCellFormula: " + ex.Message, logFilePath);
                            //double value = xlRange.Cells[i, 3].Value2;
                            //VariableFormulaValue = value.ToString("0.######");
                        }
                    }

                    if (FeatureFormulaValue != null && FeatureFormulaValue.Equals("") == false)
                    {
                        if (FeatureFormulaValue.StartsWith("=") == true)
                        {
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                            if (PartName != null && PartName.Equals("") == false && FeatureSystemName != null && FeatureSystemName.Equals("") == false)
                            {
                                String StitchFeatureName = PartName + "_" + FeatureSystemName;
                                FeatureFormulaDictionary.Add(StitchFeatureName, FeatureFormulaValue);
                            }
                        }

                    }


                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }


            }
            Marshal.ReleaseComObject(xlRange);
            xlRange = null;


            return FeatureFormulaDictionary;
        }


        private static Dictionary<String, String> ReadFormulaInFeaturesTabCorrect(string SheetName, string logFilePath, _Worksheet OldxlWorkSheet)
        {
            Dictionary<String, String> FeatureFormulaDictionary = new Dictionary<string, string>();
            //Utlity.Log("ReadInputDataFormula: ", logFilePath);

            try
            {
                if (OldxlWorkSheet == null) return null;
            }
            catch (Exception ex)
            {
                Utlity.Log("OldxlWorkSheet is NULL..." + ex.Message, logFilePath);
                return null;
            }

            Microsoft.Office.Interop.Excel.Range xlRange = OldxlWorkSheet.UsedRange;

            object[,] values = (object[,])xlRange.Formula;
            object[,] values1 = (object[,])xlRange.Value2;

            if (xlRange == null) return null;

            if (values == null)
            {
                Utlity.Log("ReadFormulaInFeaturesTabCorrect... No Formula in the Old Work sheet...", logFilePath);
                return null;
            }

            if (values1 == null)
            {
                Utlity.Log("ReadFormulaInFeaturesTabCorrect... No Values in the Old Work sheet...", logFilePath);
                return null;
            }

            //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
            //for (int i = 1; i <= xlRange.Rows.Count; i++)
            String mergePartName = "";
            for (int i = 1; i <= values.GetLength(0); i++) // Iterate on the Values, then Take the Formula
            {
                if (i == 1)
                    continue;
                try
                {
                    //Utlity.Log("Iteration: " + i.ToString(), logFilePath);
                    String PartName = "";
                    String FeatureSystemName = "";
                    String FeatureFormulaValue = "";

                    //if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                    if (values1[i, 1] != null)
                    {
                        try
                        {

                            //PartName = xlRange.Cells[i, 1].Value2;
                            PartName = Convert.ToString(values1[i, 1]);
                            mergePartName = PartName;
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("PartName: " + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        //Utlity.Log("PartName not Found.. : Setting to Top Level Part Name" + i, logFilePath);
                        PartName = mergePartName;
                        //continue;
                    }

                    //if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                    if (values1[i, 3] != null)
                    {
                        try
                        {
                            FeatureSystemName = Convert.ToString(values1[i, 3]);

                            //FeatureSystemName = xlRange.Cells[i, 3].Value2;
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("FeatureSystemName: " + ex.Message, logFilePath);
                        }
                    }
                    else
                    {
                        Utlity.Log("FeatureSystemName : " + i, logFilePath);
                        continue;
                    }

                    //if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                    if (values[i, 8] != null)
                    {
                        try
                        {
                            FeatureFormulaValue = Convert.ToString(values[i, 8]);
                            //FeatureFormulaValue = getCellFormula(xlRange, i, 8);
                            //VariableFormulaValue = getCellValue(xlRange, i, 3);
                            //VariableFormulaValue = xlRange.Cells[i, 3].Value2;
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("getCellFormula: " + ex.Message, logFilePath);
                            //double value = xlRange.Cells[i, 3].Value2;
                            //VariableFormulaValue = value.ToString("0.######");
                        }
                    }

                    if (FeatureFormulaValue != null && FeatureFormulaValue.Equals("") == false)
                    {
                        if (FeatureFormulaValue.StartsWith("=") == true)
                        {
                            //Utlity.Log("VariableName: " + VariableName, logFilePath);
                            //Utlity.Log("VariableFormulaValue: " + VariableFormulaValue, logFilePath);
                            if (PartName != null && PartName.Equals("") == false && FeatureSystemName != null && FeatureSystemName.Equals("") == false)
                            {
                                String StitchFeatureName = PartName + "_" + FeatureSystemName;
                                FeatureFormulaDictionary.Add(StitchFeatureName, FeatureFormulaValue);
                            }
                        }

                    }


                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }


            }
            Marshal.ReleaseComObject(xlRange);
            xlRange = null;


            return FeatureFormulaDictionary;
        }

        // New Sheet Processing - Update the Data from Old Sheet to the new sheet
        private static void UpdateFormulaInNewSheet(Microsoft.Office.Interop.Excel._Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook,
            Worksheet sheet, Dictionary<string, string> VariableNameFormulaDictionary, string logFilePath)
        {
            if (sheet == null) return;
            Utlity.Log(sheet.Name, logFilePath);
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            if (xlRange == null) return;
            xlApp.DisplayFormulaBar = true;

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
                            String Formula = "";
                            VariableNameFormulaDictionary.TryGetValue(name, out Formula);

                            if (Formula != null && Formula.Equals("") == false)
                            {

                                // SET VALUE
                                //xlRange.Cells[i, 3].Value2 = Formula;
                                //xlRange.Cells[i, 3].Formula = "'" + Formula;
                                //xlRange.Cells[i, 3].FormulaR1C1 = FormulaR1C1;
                                //xlRange.Cells[i, 3].Value2 = "'" + FormulaR1C1;
                                //xlRange.Cells[i, 3].Formula=  xlApp.ConvertFormula(FormulaR1C1, XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1, XlReferenceType.xlAbsolute, xlRange.Cells[i, 3]);
                                //xlRange.Cells[i, 3].FormulaLocal = FormulaLocal;                                        
                                //Utlity.Log("Formula: " + Formula, logFilePath);
                                // xlRange.Cells[i, 3].Value2 = Formula;
                                //xlRange.Cells[i, 3].Formula = Formula;

                                //xlRange.Cells[i, 3].Formula = "\"" + Formula + "\"";
                                //Utlity.Log("Value: " + xlRange.Cells[i, 3].Value2, logFilePath);
                                //xlRange.Cells[i, 3].Formula = "=SUM(1,3)";
                                // _NewRange.Formula = xlApp.ConvertFormula(Formula, XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1, XlReferenceType.xlAbsolute, xlRange.Cells[i, 3]);

                                Utlity.Log("name: " + name, logFilePath);
                                Microsoft.Office.Interop.Excel.Range _NewRange = xlRange.Cells[i, 3];
                                _NewRange.Select();
                                _NewRange.Formula = Formula;
                                _NewRange.FormulaHidden = false;
                                xlWorkSheet.Unprotect();

                                Marshal.ReleaseComObject(_NewRange);
                                _NewRange = null;


                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("UpdateFormulaInNewSheet: " + ex.Message, logFilePath);
                    }
                }

            }

            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;

                // Formula
                if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                {
                    Utlity.Log("Formula: " + xlRange.Cells[i, 3].Formula, logFilePath);
                }
            }


            NewxlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;
        }


        private static void UpdateFormulaInNewSheetCorrect(Microsoft.Office.Interop.Excel._Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook,
            Worksheet sheet, Dictionary<string, string> VariableNameFormulaDictionary, string logFilePath)
        {

            if (sheet == null) return;
            Utlity.Log("UpdateFormulaInNewSheetCorrect; " + sheet.Name, logFilePath);
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            object[,] values = (object[,])xlRange.Value2;



            if (xlRange == null) return;
            xlApp.DisplayFormulaBar = true;


            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (i == 1)
                    continue;

                // NAME
                //if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                if (values[i, 1] != null)
                {
                    try
                    {
                        String name = Convert.ToString(values[i, 1]);

                        if (name != null && name.Equals("") == false)
                        {
                            String Formula = "";
                            VariableNameFormulaDictionary.TryGetValue(name, out Formula);

                            if (Formula != null && Formula.Equals("") == false)
                            {
                                //Utlity.Log("name: " + name, logFilePath);
                                //Utlity.Log("Formula: " + Formula, logFilePath);
                                //Microsoft.Office.Interop.Excel.Range _NewRange = xlRange.Cells[i, 3];
                                //_NewRange.Select();
                                //_NewRange.Formula = Formula;

                                //_NewRange.FormulaHidden = false;
                                //xlWorkSheet.Unprotect();

                                //Marshal.ReleaseComObject(_NewRange);
                                //_NewRange = null;

                                values[i, 3] = Formula;
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("UpdateFormulaInNewSheet: " + ex.Message, logFilePath);
                    }
                }

            }


            int rows = values.GetLength(0);
            int columns = values.GetLength(1);

            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.get_Range("A1", Type.Missing);
            range = range.get_Resize(rows, columns);
            // Assign the Array to the Range in one shot:
            range.set_Value(Type.Missing, values);
            Marshal.ReleaseComObject(range);
            range = null;

            //xlRange.set_Value(Type.Missing, values);
            //xlRange.Formula = values;
            //xlRange.FormulaHidden = false;
            //xlWorkSheet.Unprotect();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

        }

        private static void UpdateFormulaInFeatureSheetCorrect(Microsoft.Office.Interop.Excel._Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook,
            Worksheet sheet, Dictionary<string, string> FeatureFormulaDictionary, string logFilePath)
        {
            if (sheet == null) return;
            if (FeatureFormulaDictionary == null || FeatureFormulaDictionary.Count == 0) return;

            //Utlity.Log(sheet.Name, logFilePath);
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            object[,] values = (object[,])xlRange.Value2;



            if (xlRange == null) return;
            xlApp.DisplayFormulaBar = true;

            String mergedPartName = "";
            String FeatureSystemName = "";
            String StitchedName = "";
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (i == 1)
                    continue;

                String PartName = "";
                if (values[i, 1] != null)
                {
                    PartName = Convert.ToString(values[i, 1]);
                    mergedPartName = PartName;
                }


                if (values[i, 3] != null)
                {
                    FeatureSystemName = Convert.ToString(values[i, 3]);
                }

                StitchedName = mergedPartName + "_" + FeatureSystemName;

                if (StitchedName != null && StitchedName.Equals("") == false)
                {
                    String Formula = "";
                    FeatureFormulaDictionary.TryGetValue(StitchedName, out Formula);

                    if (Formula != null && Formula.Equals("") == false)
                    {
                        values[i, 8] = Formula;

                    }
                }

            }

            int rows = values.GetLength(0);
            int columns = values.GetLength(1);

            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.get_Range("A1", Type.Missing);
            range = range.get_Resize(rows, columns);
            // Assign the Array to the Range in one shot:
            range.set_Value(Type.Missing, values);
            Marshal.ReleaseComObject(range);
            range = null;

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;


        }



        private static void UpdateFormulaInFeatureSheet(Microsoft.Office.Interop.Excel._Application xlApp,
            Microsoft.Office.Interop.Excel.Workbook NewxlWorkbook,
            Worksheet sheet, Dictionary<string, string> FeatureFormulaDictionary, string logFilePath)
        {
            if (sheet == null) return;
            //Utlity.Log(sheet.Name, logFilePath);
            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            if (xlRange == null) return;
            xlApp.DisplayFormulaBar = true;

            //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;

                // NAME & Feature System Name
                if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null
                    && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                {
                    try
                    {
                        String PartName = xlRange.Cells[i, 1].Value2;
                        String FeatureSystemName = xlRange.Cells[i, 3].Value2;
                        String StitchedName = PartName + "_" + FeatureSystemName;

                        if (StitchedName != null && StitchedName.Equals("") == false)
                        {
                            String Formula = "";
                            FeatureFormulaDictionary.TryGetValue(StitchedName, out Formula);

                            if (Formula != null && Formula.Equals("") == false)
                            {
                                //Utlity.Log("StitchedName: " + StitchedName, logFilePath);
                                // suppression Enabled -- Formula is Supported.
                                Microsoft.Office.Interop.Excel.Range _NewRange = xlRange.Cells[i, 8];
                                _NewRange.Select();
                                _NewRange.Formula = Formula;
                                _NewRange.FormulaHidden = false;
                                xlWorkSheet.Unprotect();

                                Marshal.ReleaseComObject(_NewRange);
                                _NewRange = null;


                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("UpdateFormulaInFeatureSheet: " + ex.Message, logFilePath);
                    }
                }

            }

            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i == 1)
                    continue;

                // Formula
                if (xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8] != null)
                {
                    Utlity.Log("Formula: " + xlRange.Cells[i, 8].Formula, logFilePath);
                }
            }


            NewxlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;
        }


        //Function written by Mani to copy images from old excel to new excel
        static void CopyImageFromOneExceltoAnother(String SourceXLFilePath, String DestXLFilePath, string logFilePath)
        {

            Utility.Log("CopyImageFromOneExceltoAnother is invoked", logFilePath);

            Microsoft.Office.Interop.Excel._Application sourcexlApp = new Microsoft.Office.Interop.Excel.Application();
            sourcexlApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel._Application destxlApp = new Microsoft.Office.Interop.Excel.Application();
            destxlApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Workbook sourcexlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks sourcexlworkbooks = null;
            Microsoft.Office.Interop.Excel.Workbook destxlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks destxlworkbooks = null;


            FileInfo sourcefile = new FileInfo(SourceXLFilePath);
            Utility.Log("SourceXLFilePath " + SourceXLFilePath, logFilePath);
            FileInfo destfile = new FileInfo(DestXLFilePath);
            Utility.Log("DestXLFilePath " + DestXLFilePath, logFilePath);

            if (sourcefile.Exists == true) // && destfile.Exists == true)
            {
                Utility.Log("sourcefile exists", logFilePath);
                sourcexlworkbooks = sourcexlApp.Workbooks;
                destxlworkbooks = destxlApp.Workbooks;

                try
                {
                    sourcexlWorkbook = sourcexlworkbooks.Open(SourceXLFilePath);
                    
                    destxlWorkbook = destxlworkbooks.Open(DestXLFilePath);
                    

                    try
                    {
                        if (sourcexlWorkbook.MultiUserEditing == true)
                        {
                            Utility.Log("Excel sharing is enabled in the workbook. Getting exclusive access", logFilePath);
                            sourcexlWorkbook.ExclusiveAccess();
                        }
                        else
                            Utility.Log("Excel sharing is not enabled in the workbook. Continuing", logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log(ex.ToString(), logFilePath);
                    }

                    try
                    {
                        if (destxlWorkbook.MultiUserEditing == true)
                        {
                            Utility.Log("Excel sharing is enabled in the workbook. Getting exclusive access", logFilePath);
                            destxlWorkbook.ExclusiveAccess();
                        }
                        else
                            Utility.Log("Excel sharing is not enabled in the workbook. Continuing", logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log(ex.ToString(), logFilePath);
                    }
                    //  Excel.Workbook destxlWorkBook = destxlApp.Workbooks.Open(DestXLFilePath, 0, false);

                    Microsoft.Office.Interop.Excel.Sheets sourcesheets = sourcexlWorkbook.Worksheets;
                    Microsoft.Office.Interop.Excel.Sheets destsheets = destxlWorkbook.Worksheets;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet sourcesheet in sourcesheets)
                    {
                        if (sourcesheet.Name.StartsWith("Input_Data", StringComparison.OrdinalIgnoreCase) == true || sourcesheet.Name.StartsWith("FEATURES", StringComparison.OrdinalIgnoreCase) == true
                            || sourcesheet.Name.StartsWith("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            Utility.Log("Skipping Custom LTC Sheet: " + sourcesheet.Name, logFilePath);
                            Marshal.ReleaseComObject(sourcesheet);
                            continue;
                        }

                        try
                        {
                            Microsoft.Office.Interop.Excel.Worksheet destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destsheets[sourcesheet.Name];

                            foreach (Microsoft.Office.Interop.Excel.Picture pic in sourcesheet.Pictures())
                            {
                                pic.CopyPicture();

                                int startCol = pic.TopLeftCell.Column;
                                int startRow = pic.TopLeftCell.Row;
                                int endCol = pic.BottomRightCell.Column;
                                int endRow = pic.BottomRightCell.Row;

                                Microsoft.Office.Interop.Excel.Range range1 = destworkSheet.Range[destworkSheet.Cells[startRow , startCol ], destworkSheet.Cells[endRow , endCol ]];
                                destworkSheet.Application.Goto(range1);
                                destworkSheet.Paste();

                                Marshal.ReleaseComObject(range1);
                                range1 = null;

                            }

                            Marshal.ReleaseComObject(destworkSheet);
                            destworkSheet = null;
                        }
                        catch (Exception ex1)
                        {
                            Utility.Log(ex1.Message, logFilePath);
                        }
                        Marshal.ReleaseComObject(sourcesheet);
                    }
                    sourcexlApp.Visible = false;
                    sourcexlApp.UserControl = false;
                    sourcexlWorkbook.Close(true);

                    destxlWorkbook.Save();
                    destxlApp.Visible = false;
                    destxlApp.UserControl = false;
                    destxlWorkbook.Close(true);

                    Marshal.ReleaseComObject(sourcesheets);
                    sourcesheets = null;
                    Marshal.ReleaseComObject(destsheets);
                    destsheets = null;

                    Marshal.ReleaseComObject(sourcexlWorkbook);
                    sourcexlWorkbook = null;
                    Marshal.ReleaseComObject(destxlWorkbook);
                    destxlWorkbook = null;

                    Marshal.ReleaseComObject(sourcexlworkbooks);
                    sourcexlworkbooks = null;
                    Marshal.ReleaseComObject(destxlworkbooks);
                    destxlworkbooks = null;

                    sourcexlApp.DisplayAlerts = true;
                    sourcexlApp.Quit();
                    destxlApp.DisplayAlerts = true;
                    destxlApp.Quit();

                    Marshal.ReleaseComObject(sourcexlApp);
                    sourcexlApp = null;
                    Marshal.ReleaseComObject(destxlApp);
                    destxlApp = null;
                }
                catch (Exception ex)
                {
                   Utility.Log(ex.Message, logFilePath);
                }
            }
            else
            {
                Console.WriteLine("file(s) does not Exist: ");
            }

            Utility.Log("CopyImageFromOneExceltoAnother completed", logFilePath);
        }

    }
}
