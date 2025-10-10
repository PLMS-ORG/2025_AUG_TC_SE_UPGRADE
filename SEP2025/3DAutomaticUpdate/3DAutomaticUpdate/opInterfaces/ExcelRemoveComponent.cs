using _3DAutomaticUpdate.controller;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate.opInterfaces
{
    class ExcelRemoveComponent
    {

        /**
       * 29 - OCT, 
       * 1) Designer Removes the Variable Part in Solid Edge
       * 2) Designer Calls SyncTE function.
       * 3) Variable Part Sheets Are Removed in Excel 
       **/

        public static void RemoveSheet(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            List<String> MasterAssemblyList = MasterAssemblyReader.getComponents();

            if (MasterAssemblyList == null || MasterAssemblyList.Count == 0)
            {
                Utility.Log("RemoveSheet: " + "Remove Component List is Empty", logFilePath);
                return;
            }

            //Utility.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utility.Log("xlApp is NULL", logFilePath);
                return;

            }
            if (xlWorkbook == null)
            {
                Utility.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
            String ltcCustomSheetName = Utility.LTCCustomSheetName;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            Utility.Log("Removing Variable Sheets, If User does Not Need It...", logFilePath);
            List<Microsoft.Office.Interop.Excel._Worksheet> sheetList = new List<_Worksheet>();
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                //Utility.Log(sheet.Name, logFilePath);
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        //Utility.Log("Skipping Custom LTC Sheet: " + sheet.Name, logFilePath);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                if (MasterAssemblyList.Contains(sheet.Name) == false)
                {
                    Utility.Log("Going to Delete Sheet: " + sheet.Name, logFilePath);
                    sheetList.Add(sheet);
                }
                else
                {
                    Marshal.ReleaseComObject(sheet);
                }
            }

            try
            {
                xlApp.DisplayAlerts = false;
                for (int i = 0; i < sheetList.Count; i++)
                {
                    Utility.Log("Going to Delete Sheet: " + sheetList[i].Name, logFilePath);
                    sheetList[i].Select();
                    sheetList[i].Activate();
                    sheetList[i].Delete();
                    Marshal.ReleaseComObject(sheetList[i]);
                    sheetList[i] = null;
                }
                xlWorkbook.Save();
                xlApp.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                Utility.Log("Deleting Sheets - Exception: " + ex.Message, logFilePath);
                Utility.Log("Deleting Sheets - Exception: " + ex.StackTrace, logFilePath);
            }

            sheetList.Clear();

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //xlApp.Visible = true;         
           
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

        }


        public static void RemoveComponentsInFeatureTAB(Microsoft.Office.Interop.Excel.Application xlApp, Microsoft.Office.Interop.Excel.Workbook xlWorkbook, String logFilePath)
        {
            List<String> MasterAssemblyList = MasterAssemblyReader.getComponents();

            if (MasterAssemblyList == null || MasterAssemblyList.Count == 0)
            {
                Utility.Log("RemoveComponentsInFeatureTAB: " + "Remove Component List is Empty", logFilePath);
                return;
            }

            //Utility.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utility.Log("xlApp is NULL", logFilePath);
                return;

            }
            if (xlWorkbook == null)
            {
                Utility.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
            String ltcCustomSheetName = Utility.LTCCustomSheetName;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            Utility.Log("Removing Variable Parts In Feature Tab, If User does Not Need It...", logFilePath);

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {

                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                    //Utility.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = xlRange.Rows.Count; i > 1; i--)
                    {
                        if (i == 1)
                            continue;

                        if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                        {
                            String componentName = xlRange.Cells[i, 6].Value2;

                            if (MasterAssemblyList.Contains(componentName) == false)
                            {
                                //Utility.Log("componentName: " + componentName, logFilePath);
                                Range range = sheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                                //setting the range for deleting the rows
                                range.EntireRow.Delete(XlDirection.xlUp);
                                Marshal.ReleaseComObject(range);
                                range = null;

                            }
                        }



                    }
                    Marshal.ReleaseComObject(xlRange);
                    xlRange = null;

                }

                Marshal.ReleaseComObject(sheet);

            }

            Marshal.ReleaseComObject(sheets);
            sheets = null;

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //xlApp.Visible = true;         

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

        }


    }
}
