using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Creo_TC_Live_Integration.TcDataManagement;
using DemoAddInTC.se;
using DemoAddInTC.utils;

using Microsoft.Office.Interop.Excel;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;

namespace DemoAddInTC.controller
{
    class SanitizeXL_PostUpload_Logic
    {
        public static Dictionary<String, String> objectNameDictionary = new Dictionary<String, String>();
        public static List<String> itemIDCollection = new List<string>();
        static SolidEdgeFramework.SolidEdgeTCE objSEEC = SEECAdaptor.getSEECObject();

        public static Dictionary<String, bool> partEnablementDictionary = new Dictionary<string, bool>();

        public static void read_all_items_in_cache(String assemblyFileName, String logFilePath)
        {
            itemIDCollection.Clear();
            objectNameDictionary.Clear();
            partEnablementDictionary.Clear();

            String bstrCachePath = SEECAdaptor.GetPDMCachePath();

            if (bstrCachePath == null || bstrCachePath == "" || bstrCachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                bstrCachePath = Path.GetDirectoryName(assemblyFileName);
            }

            traverseAssembly(assemblyFileName, logFilePath);

            foreach (String item in itemIDCollection)
            {
                Utlity.Log("traverseAssembly: " + item, logFilePath);
            }
            collect_based_on_property(itemIDCollection, logFilePath);

            foreach (KeyValuePair<String, String> kvp in objectNameDictionary)
            {
                Utility.Log("collect_based_on_property: key: Value: " + kvp.Key + ":" + kvp.Value, logFilePath);

            }

            try
            {
                read_and_sanitize_XL(bstrCachePath, assemblyFileName, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("exception in read_and_sanitize_XL function" + ex.Message, logFilePath);
                Utility.Log("Unable to read_and_sanitize_XL", logFilePath);
                return;
            }

            String Revision = Utlity.getPropertyValue("ITEM_REVISION_TYPE", logFilePath); //"P7MCPgPartRevision"
            Utility.Log("ITEM_REVISION_TYPE: " + Revision, logFilePath);
            if (Revision.Equals("ltc4_ItemRevision", StringComparison.OrdinalIgnoreCase) == true)
            {
                Utility.Log("updateSuffixProperty ", logFilePath);
                updateSuffixProperty(partEnablementDictionary, logFilePath);
            }


        }

        // update the part var type property
        public static void updateSuffixProperty(Dictionary<String, bool> LocalpartEnablementDictionary, string logFilePath)
        {
            partEnablementDictionary = LocalpartEnablementDictionary;
            //if (itemIDCollection == null && itemIDCollection.Count == 0)
            //{
            //    Utility.Log("updateSuffixProperty: itemIDCollection is Empty:", logFilePath);
            //    return;
            //}

            if (partEnablementDictionary == null && partEnablementDictionary.Count == 0)
            {
                Utility.Log("updateSuffixProperty: partEnablementDictionary is Empty:", logFilePath);
                return;
            }

            foreach (String key in partEnablementDictionary.Keys)
            {
                // key is sheet name - which is itemid.extn
                bool value;
                partEnablementDictionary.TryGetValue(key, out value);

                if (objSEEC == null)
                {
                    Utility.Log("updateSuffixProperty: objSEEC is NULL:", logFilePath);
                    return;
                }

                String bStrCachePath = "";
                objSEEC.GetPDMCachePath(out bStrCachePath);

                String documentFileName = Path.Combine(bStrCachePath, key);
                String bstrItemId = "";
                String bstrItemRev = "";

                objSEEC.GetDocumentUID(documentFileName, out bstrItemId, out bstrItemRev);

                ModelObject itemRevMO = DownloadDatasetNamedReference.getItemRevisionQuery(bstrItemId, bstrItemRev, logFilePath);

                if (itemRevMO == null)
                {

                    Utility.Log("Item Rev Model Object is NULL: " + bstrItemId, logFilePath);
                    return;
                }

                try
                {
                    if (value == true)
                    {
                        Utility.Log("ltc4_Part_Var_Type Property Value:" + "Variable", logFilePath);
                        TcAdaptor.setIRProperty(itemRevMO, "Variable", "ltc4_Part_Var_Type", logFilePath);
                    }
                    else
                    {
                        Utility.Log("ltc4_Part_Var_Type Property Value:" + "constant", logFilePath);
                        TcAdaptor.setIRProperty(itemRevMO, "constant", "ltc4_Part_Var_Type", logFilePath);
                    }

                }

                catch (Exception e)
                {
                    Utility.Log("exception in getAllDataSet function" + e.Message, logFilePath);
                }

            }
            partEnablementDictionary.Clear();

        }

        public static void collect_based_on_property(List<String> itemIDCollection, String logFilePath)
        {

            foreach (String item in itemIDCollection)
            {
                String item_ID = item.Split('~')[0];
                String rev_ID = item.Split('~')[1];

                Utility.Log("collect_based_on_property: getItemRevisionQuery:" + item_ID + "~" + rev_ID, logFilePath);
                ModelObject itemRevMO = DownloadDatasetNamedReference.getItemRevisionQuery(item_ID, rev_ID, logFilePath);

                if (itemRevMO == null)
                {

                    Utility.Log("Item Rev Model Object is NULL: " + item_ID, logFilePath);
                    return;
                }

                try
                {

                    ItemRevision itemRev = (ItemRevision)itemRevMO;
                    if (itemRev == null)
                    {
                        Utility.Log("isDataSetAvailable: " + "itemRev is null...", logFilePath);
                        return;
                    }

                    String object_name = itemRev.Object_name;

                    Utility.Log("isDataSetAvailable: object_name: " + object_name, logFilePath);
                    if (objectNameDictionary.ContainsKey(object_name) == false)
                    {
                        objectNameDictionary.Add(object_name, item_ID);
                    }

                }

                catch (Exception e)
                {
                    Utility.Log("exception in collect_based_on_property function" + e.Message, logFilePath);
                }

            }

        }

        public static void read_and_sanitize_XL(String bstrCachePath, String assemblyFileName, String logFilePath)
        {

            String XLTemplateFile = System.IO.Path.ChangeExtension(assemblyFileName, ".xlsx");
            Utility.Log("read_and_sanitize_XL.." + XLTemplateFile, logFilePath);
            if (System.IO.File.Exists(XLTemplateFile) == false)
            {
                Utility.Log("File does not Exist.." + XLTemplateFile, logFilePath);
                String[] XlFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                                             .ToArray();
                if (XlFiles.Length > 1)
                {
                    Utility.Log("read_and_sanitize_XL..could not find the exact template for assembly, " + assemblyFileName, logFilePath);
                    return;
                }
                System.IO.File.Move(XlFiles[0], XLTemplateFile);
                //XLTemplateFile = XlFiles[0];
            }

            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;



            try
            {
                // change the sheet names
                RenameExcelSheet(xlApp, XLTemplateFile, logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Could not rename sheets in Excel File.." + XLTemplateFile, logFilePath);
            }

            try
            {
                // open the FEATURES sheet and modify the partname column.
                // open the MASTER ASSEMBLY and modify -- FullName , AbsolutePath, DocNum
                SanitizeSheetInXL(xlApp, XLTemplateFile, logFilePath);
            }
            catch (Exception e)
            {
                Utility.Log("Could not SanitizeSheetInXL.." + XLTemplateFile, logFilePath);
            }

            finally
            {
                //quit and release
                if (xlApp != null) xlApp.DisplayAlerts = true;
                if (xlApp != null) xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        // Rename the excel sheet names using the based_on property.
        public static void RenameExcelSheet(Microsoft.Office.Interop.Excel._Application xlApp, string sFileName, String logFilePath)
        {
            Utlity.Log("Inside RenameExcelSheet", logFilePath);
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
            Microsoft.Office.Interop.Excel._Worksheet excelWorkSheet;

            try
            {
                excelWorkbook = xlApp.Workbooks.Open(sFileName);
                if (excelWorkbook.Sheets.Count > 0)
                {
                    String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
                    foreach (Microsoft.Office.Interop.Excel._Worksheet sheet in excelWorkbook.Sheets)
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

                        //Utlity.Log(sheet.Name, logFilePath);
                        if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true ||
                            sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            Utlity.Log("Skipping : " + sheet.Name, logFilePath);
                            Marshal.ReleaseComObject(sheet);
                            continue;
                        }

                        excelWorkSheet = sheet;
                        String oldName = excelWorkSheet.Name;
                        String name = "";
                        String extn = "";
                        if (oldName.Contains('.') == true)
                        {
                            name = oldName.Split('.')[0];
                            extn = oldName.Split('.')[1];
                        }
                        else
                        {
                            name = oldName;
                        }

                        if (objectNameDictionary.ContainsKey(oldName) == true)
                        {
                            String value = "";
                            objectNameDictionary.TryGetValue(oldName, out value);
                            //Rename the sheet
                            excelWorkSheet.Name = value + "." + extn;
                        }

                        Marshal.ReleaseComObject(excelWorkSheet);

                    }
                }

                //Save the excel 
                excelWorkbook.Save();
                //Close the excel
                excelWorkbook.Close();



                Marshal.ReleaseComObject(excelWorkbook);
                excelWorkbook = null;
            }

            catch (Exception ex)
            {
                Utility.Log("Export Excel Failed: " + ex.Message, logFilePath);
            }
        }

        public static bool SanitizeSheetInXL(Microsoft.Office.Interop.Excel._Application xlApp,
            String outputXLfileName,
          String logFilePath)
        {
            //Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            Utlity.Log("Inside SanitizeSheetInXL", logFilePath);

            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            FileInfo f = new FileInfo(outputXLfileName);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                Utlity.Log("File Already Exists,", logFilePath);
                //xlWorkbook = xlApp.Workbooks.Open(outputXLfileName);
                try
                {
                    xlWorkbook = xlApp.Workbooks.Open(outputXLfileName);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xlWorkbook = xlApp.Workbooks.Open(outputXLfileName, CorruptLoad: 1);
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.MessageBox.Show(ex1.Message);
                    }

                }
            }
            else
            {
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;
                return false;
            }
            if (xlWorkbook == null)
            {
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;
                Utlity.Log("xlWorkBook is NULL", logFilePath);
                return false;
            }
            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName(logFilePath);
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

                Utlity.Log(sheet.Name, logFilePath);

                if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true ||
                    sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                    int fullNameColNum = 0, absolutePathColNum = 0, fullNameColNumII = 0;
                    int partNameColNum = 0, partNameColNumII = 0;


                    if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        //Finding Fullname column and absolute path column number
                        Utility.Log("Attempting to find Fullname column and absolute path column number in sheet " + sheet.Name, logFilePath);
                        for (int r = 1; r <= xlRange.Rows.Count; r++)
                        {
                            for (int c = 1; c <= xlRange.Columns.Count; c++)
                            {
                                try
                                {

                                    string cellValue = null;
                                    cellValue = sheet.Cells[r, c].Value;
                                    if (cellValue != null)
                                    {
                                        Utility.Log("Found cell value is " + cellValue, logFilePath);
                                        if (cellValue.ToLower().Equals("fullname"))
                                        {
                                            if (fullNameColNum == 0)
                                                fullNameColNum = c;
                                            else
                                                fullNameColNumII = c;
                                        }
                                        else if (cellValue.ToLower().Equals("absolutepath"))
                                            absolutePathColNum = c;
                                    }

                                    //if (fullNameColNum != 0 & absolutePathColNum != 0)
                                    //    goto furtherProcesses;
                                }
                                catch (Exception ex)
                                {
                                    Utility.Log("SanitizeSheetInXL : Exception " + ex.ToString(), logFilePath);
                                    continue;
                                }

                            }
                        }

                    furtherProcesses: ;
                        Utility.Log("Full Name column number is " + fullNameColNum.ToString(), logFilePath);
                        Utility.Log("Full NameII column number is " + fullNameColNumII.ToString(), logFilePath);
                        Utility.Log("Absolute Path column number is " + absolutePathColNum.ToString(), logFilePath);
                        if (fullNameColNum != 0 & absolutePathColNum != 0)
                            Utility.Log("Both required columns have been found. Continuing", logFilePath);
                        else
                        {
                            Utility.Log("Cannot sanitize excel as the required columns were not found", logFilePath);
                            return false;
                        }
                    }
                    else
                    {
                        Utility.Log("Attempting to find Partname column number in " + sheet.Name, logFilePath);

                        for (int r = 1; r <= xlRange.Rows.Count; r++)
                        {
                            for (int c = 1; c <= xlRange.Columns.Count; c++)
                            {
                                try
                                {

                                    string cellValue = null;
                                    cellValue = sheet.Cells[r, c].Value;
                                    if (cellValue != null)
                                    {
                                        Utility.Log("Found cell value is " + cellValue, logFilePath);
                                        if (cellValue.ToLower().Equals("partname"))
                                        {
                                            if (partNameColNum == 0)
                                                partNameColNum = c;
                                            else
                                                partNameColNumII = c;
                                        }
                                    }

                                    //if (partNameColNum != 0)
                                    //    goto furtherProcesses;
                                }
                                catch (Exception ex)
                                {
                                    Utility.Log("SanitizeSheetInXL : Exception " + ex.ToString(), logFilePath);
                                    continue;
                                }
                            }
                        }

                    furtherProcesses: ;
                        Utility.Log("PartName column number is " + partNameColNum.ToString(), logFilePath);
                        Utility.Log("partNameII column number is " + partNameColNumII.ToString(), logFilePath);
                        if (partNameColNum != 0)
                            Utility.Log("Required column has been found. Continuing", logFilePath);
                        else
                        {
                            Utility.Log("Cannot sanitize excel as the required column was not found", logFilePath);
                            return false;
                        }
                    }


                    try
                    {
                        Utlity.Log("sheet.Name" + ":::" + sheet.Name, logFilePath);
                        foreach (string key in objectNameDictionary.Keys)
                        {
                            //WriteFeatureSheet(xlWorkbook, sheet, sheet.Name, logFilePath);
                            String value = "";

                            String name = "";
                            String extn = "";
                            if (key.Contains('.') == true)
                            {
                                if (key.Split('.').Length == 2)
                                {
                                    name = key.Split('.')[0];
                                    extn = key.Split('.')[1];
                                }
                                else
                                {
                                    string[] split = key.Split('.');
                                    int splitLength = split.Length;
                                    int extnPos = splitLength - 1;
                                    for (int i = 0; i < splitLength; i++)
                                    {
                                        if (i == 0)
                                            name = split[i];
                                        else if (i != extnPos)
                                            name = name + "." + split[i];
                                        else if (i == extnPos)
                                            extn = split[i];
                                    }
                                }
                            }
                            else
                            {
                                name = key;
                            }


                            objectNameDictionary.TryGetValue(key, out value);
                            Utility.Log("Key : " + key + " Value : " + value, logFilePath);


                            if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                //Iterating through fullname column and absolute path column to replace text
                                Utility.Log("Iterating through absolute path column to replace text", logFilePath);
                                for (int r = 1; r <= xlRange.Rows.Count; r++)
                                {
                                    for (int c = 1; c <= xlRange.Columns.Count; c++)
                                    {
                                        try
                                        {
                                            if (c == absolutePathColNum)
                                            {

                                                string cellValue = null;
                                                cellValue = sheet.Cells[r, c].Value;
                                                if (cellValue != null)
                                                {
                                                    Utility.Log("Found cell value is " + cellValue, logFilePath);
                                                    if (cellValue.Contains("\\") & cellValue.ToLower().Contains(key.ToLower()))
                                                    {
                                                        Utility.Log("Cell value contains \\ & " + key, logFilePath);
                                                        string[] splitValue = cellValue.Split('\\');
                                                        Utility.Log("The last cell value after splitting is " + splitValue.Last(), logFilePath);
                                                        if (splitValue.Last().ToLower().Equals(key.ToLower()))
                                                        {
                                                            Utlity.Log("Key " + key + " and last cell value " + splitValue.Last() + " match", logFilePath);
                                                            Array.Resize(ref splitValue, splitValue.Length - 1);
                                                            string newValue = Path.Combine(splitValue);
                                                            newValue = newValue + "\\" + value + "." + extn;
                                                            Utlity.Log(cellValue + " to be replaced by " + newValue, logFilePath);
                                                            sheet.Cells[r, c].Value = newValue;
                                                        }
                                                    }

                                                }

                                            }
                                            //else if (c == fullNameColNum)
                                            //{

                                            //    string cellValue = null;
                                            //    cellValue = sheet.Cells[r, c].Value;
                                            //    if (cellValue != null)
                                            //    {
                                            //        Utility.Log("Found cell value is " + cellValue, logFilePath);
                                            //        if (key.ToLower().Equals(cellValue.ToLower()) || key.ToLower().Equals(cellValue.ToLower() + "." + extn.ToLower()))
                                            //        {
                                            //            //string newValue = cellValue.Replace(key, value + "." + extn);
                                            //            Utlity.Log(cellValue + " to be replaced by " + value + "." + extn, logFilePath);
                                            //            sheet.Cells[r, c].Value = value + "." + extn;
                                            //        }
                                            //    }

                                            //}
                                        }
                                        catch (Exception ex)
                                        {
                                            Utility.Log("SanitizeSheetInXL : Exception " + ex.ToString(), logFilePath);
                                            continue;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                //Iterating through Partname column to replace text
                                Utility.Log("Iterating through Partname column to replace text", logFilePath);
                                for (int r = 1; r <= xlRange.Rows.Count; r++)
                                {
                                    for (int c = 1; c <= xlRange.Columns.Count; c++)
                                    {
                                        try
                                        {
                                            if (c == partNameColNum)
                                            {

                                                string cellValue = null;
                                                cellValue = sheet.Cells[r, c].Value;
                                                if (cellValue != null)
                                                {
                                                    Utility.Log("Found cell value is " + cellValue, logFilePath);
                                                    if (key.ToLower().Equals(cellValue.ToLower()) || key.ToLower().Equals(cellValue.ToLower() + "." + extn.ToLower()))
                                                    {
                                                        //string newValue = cellValue.Replace(key, value + "." + extn);
                                                        Utlity.Log(cellValue + " to be replaced by " + value + "." + extn, logFilePath);
                                                        sheet.Cells[r, c].Value = value + "." + extn;
                                                        if (partNameColNumII != 0)
                                                            sheet.Cells[r, partNameColNumII].Value = value + "." + extn;
                                                    }
                                                }

                                            }
                                            if (partNameColNumII != 0)
                                            {
                                                if (c == partNameColNumII)
                                                {
                                                    string cellValue = null;
                                                    cellValue = sheet.Cells[r, c].Value;
                                                    if (cellValue != null)
                                                    {
                                                        Utility.Log("Found cell value is " + cellValue, logFilePath);
                                                        if (key.ToLower().Equals(cellValue.ToLower()) || key.ToLower().Equals(cellValue.ToLower() + "." + extn.ToLower()))
                                                        {
                                                            //string newValue = cellValue.Replace(key, value + "." + extn);
                                                            Utlity.Log(cellValue + " to be replaced by " + value + "." + extn, logFilePath);
                                                            sheet.Cells[r, c].Value = value + "." + extn;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Utility.Log("SanitizeSheetInXL : Exception " + ex.ToString(), logFilePath);
                                            continue;
                                        }

                                    }
                                }
                            }

                        }



                        //FindAndReplace(xlWorkbook, sheet, key, value + "." + extn, logFilePath);

                        if (sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            Utility.Log("Iterating to replace fullname column text", logFilePath);
                            for (int r = 2; r <= xlRange.Rows.Count; r++)
                            {
                                try
                                {
                                    string apValue = sheet.Cells[r, absolutePathColNum].Value;
                                    string fnValue = apValue.Split('\\').Last();
                                    Utility.Log("Replacing " + apValue + " with " + fnValue, logFilePath);
                                    sheet.Cells[r, fullNameColNum].Value = fnValue;
                                    if (fullNameColNumII != 0)
                                        sheet.Cells[r, fullNameColNumII].Value = apValue;
                                }
                                catch (Exception ex)
                                {
                                    Utility.Log("Fullname exception " + ex.ToString(), logFilePath);
                                    continue;
                                }

                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        //Marshal.ReleaseComObject(xlWorkbook);
                        //xlWorkbook = null;
                        //Marshal.ReleaseComObject(workbooks);
                        //workbooks = null;
                        Marshal.ReleaseComObject(sheet);
                        Utlity.Log("SanitizeSheetInXL: Exception " + ex.Message, logFilePath);
                        continue;
                    }
                }
                else
                {
                    try
                    {
                        WriteSheet(xlWorkbook, sheet, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("WriteSheet:" + ex.Message, logFilePath);
                        Utlity.Log("WriteSheet:" + ex.StackTrace, logFilePath);
                        Utlity.Log("WriteSheet:" + sheet.Name, logFilePath);
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

                }


                Marshal.ReleaseComObject(sheet);
            }

            xlWorkbook.Close(true);

            Marshal.ReleaseComObject(sheets);
            sheets = null;


            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;


            Marshal.ReleaseComObject(workbooks);
            workbooks = null;



            return true;
        }


        public static void FindAndReplace(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel.Worksheet sheet,
            String find, String replace, String logFilePath)
        {
            Utility.Log("FindAndReplace: find-" + find, logFilePath);
            Utility.Log("FindAndReplace: replace-" + replace, logFilePath);

            if (sheet == null)
            {
                Utlity.Log("FindAndReplace: sheet is Empty", logFilePath);
                return;
            }


            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;





            // call the replace method to replace instances. 
            bool success = (bool)xlRange.Replace(
                find,
                replace,
                XlLookAt.xlPart,
                XlSearchOrder.xlByRows,
                false, Type.Missing, Type.Missing, Type.Missing);

            xlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;



        }

        // another copy from SolidEdgeData1 class
        public static void traverseAssembly(String assemblyFileName, String logFilePath)
        {

            // 19/8 - Purpose - To find repetitive BomLines & Not Add into bomLineList (Object Store)
            // If Added Again into Object Store, Issues Arise in TreeView in the UI.   
            if (assemblyFileName == null || File.Exists(assemblyFileName) == false)
            {
                Utlity.Log("traverseAssembly: " + "assemblyFileName does not Exist.." + assemblyFileName, logFilePath);
                return;
            }

            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            SolidEdge.RevisionManager.Interop.Document document = null;
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;
            SolidEdge.RevisionManager.Interop.Document linkDocument = null;
            try
            {
                document = objReviseApp.Open(assemblyFileName);
                if (document == null)
                {
                    Utlity.Log("traverseAssembly: " + "Document is NULL", logFilePath);
                    return;
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("traverseAssembly: " + ex.Message, logFilePath);
                return;
            }

            String bstrItemId = "";
            String bstrItemRev = "";
            //String AbsolutePath = document.AbsolutePath;
            String AbsolutePath = document.FullName;

            objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);

            if (itemIDCollection.Contains(bstrItemId + "~" + bstrItemRev) == false)
            {
                itemIDCollection.Add(bstrItemId + "~" + bstrItemRev);
            }



            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + assemblyFileName, logFilePath);
                return;

            }

            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];
                if (linkDocument.FullName.EndsWith(".xlsx") == true)
                {
                    Utlity.Log("Skipping: " + linkDocument.FullName, logFilePath);
                    continue;
                }

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);
                //Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                //Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                //Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                //Utlity.Log("Status: " + linkDocument.Status, logFilePath);

                //AbsolutePath = linkDocument.AbsolutePath;
                AbsolutePath = linkDocument.FullName;

                objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
                Utlity.Log("AbsolutePath: " + AbsolutePath + ":::" + bstrItemId + ":::" + bstrItemRev, logFilePath);
                if (itemIDCollection.Contains(bstrItemId + "~" + bstrItemRev) == false)
                {
                    itemIDCollection.Add(bstrItemId + "~" + bstrItemRev);
                }


                //bomLineList.Add(bl);
                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument, logFilePath);

                }
            }

            SE_SESSION.killRevisionManager(logFilePath);

        }

        private static void traverseLinkDocuments(SolidEdge.RevisionManager.Interop.Document linkDocument, String logFilePath)
        {
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + linkDocument.FullName, logFilePath);
                return;
            }

            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);

                //Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                //Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                //Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                //Utlity.Log("Status: " + linkDocument.Status, logFilePath);

                String bstrItemId = "";
                String bstrItemRev = "";

                //String AbsolutePath = linkDocument.AbsolutePath;
                String AbsolutePath = linkDocument.FullName;

                objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
                //Utlity.Log("AbsolutePath: " + AbsolutePath + ":::" + bstrItemId + ":::" + bstrItemRev, logFilePath);
                if (itemIDCollection.Contains(bstrItemId + "~" + bstrItemRev) == false)
                {
                    itemIDCollection.Add(bstrItemId + "~" + bstrItemRev);
                }

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument, logFilePath);

                }
            }



        }

        private static void WriteSheet(Microsoft.Office.Interop.Excel.Workbook xlWorkbook,
            Microsoft.Office.Interop.Excel.Worksheet Sheet, String logFilePath)
        {


            Microsoft.Office.Interop.Excel._Worksheet xlWorkSheet = Sheet;
            xlWorkSheet.Activate();

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            try
            {
                for (int i = 1; i <= xlRange.Rows.Count; i++)
                //for (int i = 0; i < xlRange.Rows.Count; i++)
                {
                    if (i == 1) continue;
                    String oldName = xlRange.Cells[i, 9].Value2;
                    String name = "";
                    String extn = "";

                    if (oldName == null || oldName.Equals("") == true)
                    {
                        //Utility.Log("oldName: " + oldName, logFilePath);
                        continue;
                    }
                    if (oldName.Contains('.') == true)
                    {
                        name = oldName.Split('.')[0];
                        extn = oldName.Split('.')[1];
                    }
                    else
                    {
                        name = oldName;
                    }

                    //Utility.Log("name: " + name, logFilePath);

                    if (objectNameDictionary.ContainsKey(oldName) == true)
                    {
                        String value = "";
                        objectNameDictionary.TryGetValue(oldName, out value);
                        xlRange.Cells[i, 9].Value2 = value + "." + extn;
                    }

                }
            }
            catch (Exception ex)
            {
                Utility.Log("WriteSheet: " + ex.Message, logFilePath);
                Utility.Log("WriteSheet: " + ex.StackTrace, logFilePath);
            }

            xlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

            //Marshal.ReleaseComObject(xlWorkSheet);
            //xlWorkSheet = null;

        }

    }
}
