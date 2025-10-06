using ExcelSyncTC.controller;
using ExcelSyncTC.utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelSyncTC.opInterfaces
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
                Utlity.Log("RemoveSheet: " + "Remove Component List is Empty", logFilePath);
                return;
            }

             //Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utlity.Log("xlApp is NULL", logFilePath);
                return;

            }           
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            Utlity.Log("Removing Variable Sheets, If User does Not Need It...", logFilePath);
            List<Microsoft.Office.Interop.Excel._Worksheet> sheetList = new List<_Worksheet>();
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                //Utlity.Log(sheet.Name, logFilePath);
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
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true || sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Marshal.ReleaseComObject(sheet);
                    continue;
                }

                if (MasterAssemblyList.Contains(sheet.Name) == false)
                {
                    Utlity.Log("Going to Delete Sheet: " + sheet.Name, logFilePath);
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
                    sheetList[i].Select();
                    sheetList[i].Activate();
                    sheetList[i].Delete();
                    Marshal.ReleaseComObject(sheetList[i]);
                    sheetList[i] = null;
                }
                xlApp.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                Utlity.Log("Deleting Sheets - Exception: " + ex.Message, logFilePath);
                Utlity.Log("Deleting Sheets - Exception: " + ex.StackTrace, logFilePath);
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
                Utlity.Log("RemoveComponentsInFeatureTAB: " + "Remove Component List is Empty", logFilePath);
                return;
            }

            //Utlity.Log("----------------------------------------------------------", logFilePath);
            if (xlApp == null)
            {
                Utlity.Log("xlApp is NULL", logFilePath);
                return;

            }
            if (xlWorkbook == null)
            {
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return;
            }
            String ltcCustomSheetName = Utlity.LTCCustomSheetName;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            Utlity.Log("Removing Variable Parts In Feature Tab, If User does Not Need It...", logFilePath);
          
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {
                               
                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                     Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                    //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                     for (int i = xlRange.Rows.Count; i > 1; i--)
                     {
                         if (i == 1)
                             continue;

                         if (xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6] != null)
                         {
                             String componentName = xlRange.Cells[i, 6].Value2;

                              if (MasterAssemblyList.Contains(componentName) == false) 
                              {
                                  //Utlity.Log("componentName: " + componentName, logFilePath);
                                  Range range = sheet.get_Range("A" + i.ToString(),"A" + i.ToString());
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
