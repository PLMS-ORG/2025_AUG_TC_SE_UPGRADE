using DemoAddInTC.controller;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DemoAddInTC.opInterfaces
{
    class ExcelComponentDeltaInterface
    {
        // 27 - OCT, Feature Sync from Excel would Fail if Variable Parts are not Renamed in Features TAB.
        public static void RenameComponentDetailsInMasterAssemblyTab(String assemblyFileName, String FolderToPublish, List<String> variableParts, String Suffix,
            String logFilePath,
            String Option)
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



            if (xlFilePath == null || xlFilePath.Equals("") == true)
            {
                Utlity.Log("RenameSheetNamesForVariableParts: XLFilePath is Empty..", logFilePath);
                return;
            }

         
            if (Option.Equals("CTD", StringComparison.OrdinalIgnoreCase) == true)
            {
                UpdateBOMLineDetailsAfterCTDInMasterAssembly(xlFilePath, logFilePath, variableParts, Suffix, FolderToPublish);
            }
            else
            {
                UpdateBOMLineDetailsInMasterAssembly(xlFilePath,FolderToPublish,  logFilePath);

            }

            return;
        }

        public static void UpdateBOMLineDetailsAfterCTDInMasterAssembly(String xlFilePath,String logFilePath,
            List<String> variablePartsList,
            String Suffix,
            String FolderToPublish)
        {
            Utlity.Log("UpdateBOMLineDetailsAfterCTDInMasterAssembly: XLFilePath: " + xlFilePath, logFilePath);
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
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
                    xlWorkSheet.Activate();
                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i == 1)
                            continue;
                        String UpdatedFilePath = "";

                        

                        String NewFileName = "";
                        if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // AbsolutePath
                                filePath = xlRange.Cells[i, 3].Value2;
                                Utlity.Log("filePath: " + filePath, logFilePath);
                                String fileName = Path.GetFileName(filePath);
                                Utlity.Log("fileName: " + fileName, logFilePath);
                                if (variablePartsList.Contains(fileName) == true)
                                {
                                    NewFileName = Path.GetFileNameWithoutExtension(filePath) + Suffix ;
                                    Utlity.Log(NewFileName, logFilePath);
                                    xlRange.Cells[i, 1].Value2 = NewFileName;
                                }
                                if (NewFileName != null && NewFileName.Equals("") == false)
                                {
                                    UpdatedFilePath = Path.Combine(FolderToPublish, NewFileName + Path.GetExtension(filePath));
                                    if (UpdatedFilePath != null && UpdatedFilePath.Equals("") == false)
                                    {
                                        xlRange.Cells[i, 3].Value2 = UpdatedFilePath;
                                    }
                                }
                                else
                                {
                                    // All Parts should be Updated. Not Only Variable Parts - 30/01 - Murali
                                    String newFileName = fileName; // (No Suffix would be Added here)
                                    Utlity.Log("newFileName3: " + fileName, logFilePath);
                                    UpdatedFilePath = Path.Combine(FolderToPublish, fileName);
                                    if (UpdatedFilePath != null && UpdatedFilePath.Equals("") == false)
                                    {
                                        xlRange.Cells[i, 3].Value2 = UpdatedFilePath;
                                    }

                                }

                                Utlity.Log("UpdatedFilePath3" + UpdatedFilePath, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AbsolutePath" + ex.Message, logFilePath);
                            }
                        }


                        if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // FullName
                                filePath = xlRange.Cells[i, 7].Value2;
                                String fileName = Path.GetFileName(filePath);
                                Utlity.Log("fileName: " + fileName, logFilePath);
                                if (variablePartsList.Contains(fileName) == true)
                                {
                                    NewFileName = Path.GetFileNameWithoutExtension(filePath) + Suffix + Path.GetExtension(filePath);                                    
                                }
                                if (NewFileName != null && NewFileName.Equals("") == false)
                                {
                                    UpdatedFilePath = Path.Combine(FolderToPublish, NewFileName);
                                    if (UpdatedFilePath != null && UpdatedFilePath.Equals("") == false)
                                    {
                                        xlRange.Cells[i, 7].Value2 = UpdatedFilePath;
                                    }
                                }
                                else
                                {
                                    // All Parts should be Updated. Not Only Variable Parts - 30/01 - Murali
                                    String newFileName = fileName; // (No Suffix would be Added here)
                                    Utlity.Log("newFileName7: " + fileName, logFilePath);
                                    UpdatedFilePath = Path.Combine(FolderToPublish, fileName);
                                    if (UpdatedFilePath != null && UpdatedFilePath.Equals("") == false)
                                    {
                                        xlRange.Cells[i, 7].Value2 = UpdatedFilePath;
                                    }

                                }

                                Utlity.Log("UpdatedFilePath7" +  UpdatedFilePath, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("FullName" + ex.Message, logFilePath);
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

        // CTC, Duplicate -- In these two cases the Path Needs to be Updated in the Master Assembly TAB
        public static void UpdateBOMLineDetailsInMasterAssembly(String xlFilePath, String folderToPublish, String logFilePath)
        {
            Utlity.Log("UpdateBOMLineDetailsInMasterAssembly: XLFilePath: " + xlFilePath, logFilePath);
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
                //Utlity.Log(sheet.Name, logFilePath);
                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = sheet;
                    xlWorkSheet.Activate();
                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    //Utlity.Log(xlRange.Rows.Count.ToString(), logFilePath);
                    for (int i = 1; i <= xlRange.Rows.Count; i++)
                    {
                        if (i == 1)
                            continue;

                        if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                        {
                            String FullName = "";

                            // FullName
                            FullName = xlRange.Cells[i, 1].Value2;
                            FullName = Path.GetFileNameWithoutExtension(FullName);
                            Utlity.Log("FullName: " + FullName, logFilePath);
                            xlRange.Cells[i, 1].Value2 = FullName;
                        }

                        String UpdatedFilePath = "";
                        if (xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // AbsolutePath
                                filePath = xlRange.Cells[i, 3].Value2;
                                String fileName = Path.GetFileName(filePath);
                                UpdatedFilePath = Path.Combine(folderToPublish, fileName);
                                if (UpdatedFilePath != null && UpdatedFilePath.Equals("") == false)
                                {
                                    xlRange.Cells[i, 3].Value2 = UpdatedFilePath;
                                }

                                //Utlity.Log(fileName, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("AbsolutePath" + ex.Message, logFilePath);
                            }
                        }
                        

                        if (xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7] != null)
                        {
                            try
                            {
                                String filePath = "";
                                // FullName
                                filePath = xlRange.Cells[i, 7].Value2;

                                String fileName = Path.GetFileName(filePath);
                                UpdatedFilePath = Path.Combine(folderToPublish, fileName);
                                if (UpdatedFilePath != null && UpdatedFilePath.Equals("") == false)
                                {
                                    xlRange.Cells[i, 7].Value2 = UpdatedFilePath;
                                }

                                //Utlity.Log(filePath, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("FullName" + ex.Message, logFilePath);
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
