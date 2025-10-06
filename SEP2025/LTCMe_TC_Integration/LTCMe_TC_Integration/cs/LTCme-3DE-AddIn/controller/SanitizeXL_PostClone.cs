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
    class SanitizeXL_PostClone
    {
        public static Dictionary<String, String> basedOnDictionaryBOAKey = new Dictionary<String, String>();
        public static Dictionary<String, String> basedOnDictionaryitemIDKey = new Dictionary<String, String>();
        public static List<String> itemIDCollection = new List<string>();
        public static SolidEdgeFramework.SolidEdgeTCE objSEEC = SEECAdaptor.getSEECObject();
        public static Dictionary<String, bool> partEnablementDictionary = new Dictionary<string, bool>();


        public static void read_all_items_in_cache_SEEC(String assemblyFileName, String logFilePath)
        {
            basedOnDictionaryBOAKey.Clear();
            basedOnDictionaryitemIDKey.Clear();
            itemIDCollection.Clear();

            String bstrCachePath = SEECAdaptor.GetPDMCachePath();


            Utility.Log("read_all_items_in_cache: Traversing the assembly: " + assemblyFileName, logFilePath);
            traverseAssembly(assemblyFileName, logFilePath);

            Utility.Log("read_all_items_in_cache: Collecting based on Property: " + assemblyFileName, logFilePath);
            collect_based_on_property(itemIDCollection, logFilePath);

            Utility.Log("read_all_items_in_cache: read and Sanitize the XL: " + assemblyFileName, logFilePath); // -- 
            read_and_sanitize_XL(bstrCachePath, assemblyFileName, logFilePath);

            String xlFile = Path.ChangeExtension(assemblyFileName, ".xlsx");
            if (File.Exists(xlFile) == true)
            {
                // Murali - 25-NOV-2024 - SOA Decustomization - Start
                //Utlity.Log("Upload To TC Using SEEC", logFilePath, "TVS");
                Ribbon2d.UploadtoTCUsingSEEC(Ribbon2d.currentSETCEObject, Ribbon2d.currentSESession);
                // Murali - 25-NOV-2024 - SOA Decustomization - End
            }
            else
            {
                Utility.Log("read_all_items_in_cache: " + "no xl file is found in cache", logFilePath);
            }

        }
        public static void read_all_items_in_cache(String assemblyFileName, String logFilePath)
        {
            basedOnDictionaryBOAKey.Clear();
            basedOnDictionaryitemIDKey.Clear();
            itemIDCollection.Clear();

            String bstrCachePath = SEECAdaptor.GetPDMCachePath();

            if (bstrCachePath == null || bstrCachePath == "" || bstrCachePath.Equals("") == true)
            {
                Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                bstrCachePath = Path.GetDirectoryName(assemblyFileName);
            }

            Utility.Log("read_all_items_in_cache: Traversing the assembly: " + assemblyFileName, logFilePath);
            traverseAssembly(assemblyFileName, logFilePath);

            Utility.Log("read_all_items_in_cache: Collecting based on Property: " + assemblyFileName, logFilePath);
            collect_based_on_property(itemIDCollection, logFilePath);

            Utility.Log("read_all_items_in_cache: read and Sanitize the XL: " + assemblyFileName, logFilePath); // -- 
            read_and_sanitize_XL(bstrCachePath,assemblyFileName, logFilePath);

            String xlFile = Path.ChangeExtension(assemblyFileName, ".xlsx");
            if (File.Exists(xlFile) == true)
            {
                Utility.Log("read_all_items_in_cache: Upload the excel back to Teamcenter: " + assemblyFileName, logFilePath);
                TcAdaptor.uploadExcelToTC(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, xlFile, logFilePath);

                String Revision = Utlity.getPropertyValue("ITEM_REVISION_TYPE", logFilePath); //"P7MCPgPartRevision"
                Utility.Log("ITEM_REVISION_TYPE: " + Revision, logFilePath);
                if (Revision.Equals("ltc4_ItemRevision", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Utility.Log("updateSuffixProperty NEW 2024 ", logFilePath);
                    updateSuffixProperty(itemIDCollection, logFilePath, bstrCachePath);
                }
            }
            else
            {
                Utility.Log("read_all_items_in_cache: " + "no xl file is found in cache", logFilePath);
            }

        }

        // update the part var type property
        public static void updateSuffixProperty(List<String> itemIDCollection, string logFilePath, String bstrCachePath)
        {
            if (itemIDCollection == null && itemIDCollection.Count == 0)
            {
                Utility.Log("updateSuffixProperty: itemIDCollection is Empty:", logFilePath);
                return;
            }

            if (partEnablementDictionary == null && partEnablementDictionary.Count == 0)
            {
                Utility.Log("updateSuffixProperty: partEnablementDictionary is Empty:", logFilePath);
                return;
            }
            Utility.Log("updateSuffixProperty: partEnablementDictionary Count :" + partEnablementDictionary.Count, logFilePath);
            foreach (String key in partEnablementDictionary.Keys)
            {
                // key is sheet name - which is itemid.extn
                bool value;
                partEnablementDictionary.TryGetValue(key, out value);
                Utility.Log("updateSuffixProperty: partEnablementDictionary Key :" + key, logFilePath);
                Utility.Log("updateSuffixProperty: partEnablementDictionary Value :" + value, logFilePath);
                if (objSEEC == null)
                {
                    Utility.Log("updateSuffixProperty: objSEEC is NULL:", logFilePath);
                    return;
                }

                //String bStrCachePath = "";
                //objSEEC.GetPDMCachePath(out bStrCachePath);

                //if (bStrCachePath == null || bStrCachePath == "" || bStrCachePath.Equals("") == true)
                //{
                //    Utlity.Log("cachePath is NULL, So getting it from:  " + assemblyFileName, logFilePath);
                //    bStrCachePath = Path.GetDirectoryName(assemblyFileName);
                //}

                String documentFileName = Path.Combine(bstrCachePath, key);
                String bstrItemId = "";
                String bstrItemRev = "";

                objSEEC.GetDocumentUID(documentFileName, out bstrItemId, out bstrItemRev);

                Utility.Log("updateSuffixProperty::documentFileName: " + documentFileName, logFilePath);
                Utility.Log("updateSuffixProperty::bstrItemId: " + bstrItemId, logFilePath);
                Utility.Log("updateSuffixProperty::bstrItemRev: " + bstrItemRev, logFilePath);

                if (bstrItemId == null || bstrItemId.Equals("") == true || bstrItemRev == null ||
                    bstrItemRev.Equals("") == true)
                {
                    Utility.Log("bstrItemId is empty: " + bstrItemId, logFilePath);
                    Utility.Log("bstrItemRev is empty: " + bstrItemRev, logFilePath);
                    continue;
                }

                ModelObject itemRevMO = DownloadDatasetNamedReference.getItemRevisionQuery(bstrItemId, bstrItemRev, logFilePath);

                if (itemRevMO == null)
                {

                    Utility.Log("Item Rev Model Object is NULL: " + bstrItemId, logFilePath);
                    continue;
                }

                try
                {
                    if (value == true)
                    {
                        Utility.Log("ltc4_Part_Var_Type Property Value:" + "Variable", logFilePath);
                        TcAdaptor.setIRProperty(itemRevMO, "Variable", "ltc4_Part_Var_Type", logFilePath);
                        try
                        {
                            Utlity.Log("Suffix property from CTD : " + Ribbon2d.tcw.suffix.Text, logFilePath, "CTD");
                            TcAdaptor.setIRProperty(itemRevMO, Ribbon2d.tcw.suffix.Text, "ltc4_Suffix1", logFilePath);
                        }
                        catch (Exception)
                        {
                            Utlity.Log("Error encountered while updating suffix", logFilePath, "CTD");
                        }
                        
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



        }

        public static void collect_based_on_property(List<String>itemIDCollection,String logFilePath) 
        
        {

            foreach (String item in itemIDCollection)
            {
                String item_ID = item.Split('~')[0];
                String rev_ID = item.Split('~')[1];

                Utility.Log("collect_based_on_property: getItemRevisionQuery:" + item_ID + "~" + rev_ID, logFilePath);
               ModelObject itemRevMO = DownloadDatasetNamedReference.getItemRevisionQuery(item_ID, rev_ID,logFilePath);

                if (itemRevMO == null)
                {

                    Utility.Log("Item Rev Model Object is NULL: " + item_ID, logFilePath);
                    return;
                }

                try
                {
                    ItemRevision itemRev = (ItemRevision)itemRevMO;
                    if (itemRevMO == null)
                    {
                        Utility.Log("collect_based_on_property: " + "itemRev is null...", logFilePath);
                        return;
                    }

                    String based_on = itemRev.Based_on;
                    String based_on_actual = based_on.Split('/')[0];

                    if (basedOnDictionaryBOAKey.ContainsKey(based_on_actual) == false)
                    {
                        Utility.Log("basedOnDictionaryBOAKey: " + "based_on_actual..." + based_on_actual, logFilePath);
                        Utility.Log("basedOnDictionaryBOAKey: " + "item_ID..." + item_ID, logFilePath);
                        basedOnDictionaryBOAKey.Add(based_on_actual, item_ID);
                    }

                    if (basedOnDictionaryitemIDKey.ContainsKey(item_ID) == false)
                    {
                        Utility.Log("basedOnDictionaryitemIDKey: " + "based_on_actual..." + based_on_actual, logFilePath);
                        Utility.Log("basedOnDictionaryitemIDKey: " + "item_ID..." + item_ID, logFilePath);
                        basedOnDictionaryitemIDKey.Add(item_ID, based_on_actual);
                    }
                 
                }

                catch (Exception e)
                {
                    Utility.Log("exception in basedOnDictionaryBOAKey function" + e.Message, logFilePath);
                }

            }

        }

        public static void read_and_sanitize_XL(String bstrCachePath, String assemblyFileName, String logFilePath)
        {
            String XLTemplateFile = System.IO.Path.ChangeExtension(assemblyFileName, ".xlsx");
            String bstrItemId = "";
            String bstrItemRev = "";
            Utility.Log("read_and_sanitize_XL for : , " + XLTemplateFile, logFilePath);
        
       
                Utility.Log("Getting objSECC..." + XLTemplateFile, logFilePath);

                if (objSEEC == null)
                {
                    Utility.Log("read_and_sanitize_XL objSEEC is NULL : , " + XLTemplateFile, logFilePath);
                    return;
                }
            
            objSEEC.GetDocumentUID(assemblyFileName, out bstrItemId, out bstrItemRev);
            if (System.IO.File.Exists(XLTemplateFile) == false)
            {
                //Utility.Log("File does not Exist.." + XLTemplateFile, logFilePath);
                //String value = "";
                //if (basedOnDictionaryitemIDKey.ContainsKey(bstrItemId) == true)
                //{
                //    basedOnDictionaryitemIDKey.TryGetValue(bstrItemId, out value);
                //}
                //if (value == null || value.Equals("") == true)
                //{
                //    Utility.Log("read_and_sanitize_XL..could not find the exact template for assembly, " + assemblyFileName, logFilePath);
                //    return;
                //}
                //String newXLFile = Path.Combine(bstrCachePath, value + ".xlsx");

                String[] XlFiles = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.EndsWith("_Structured.xlsx", StringComparison.OrdinalIgnoreCase)) ||
                                                 (x.EndsWith("_Consolidated.xlsx", StringComparison.OrdinalIgnoreCase)))
                                             .ToArray();
                foreach (String xLFile in XlFiles)
                {
                    Utility.Log("Deleting xLFile: , " + xLFile, logFilePath);
                    File.Delete(xLFile);                   
                }

                String[] tempXL = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                                             .ToArray();

                if (tempXL.Length == 0)
                {
                    Utility.Log("No XL found inside the cache, cannot sanitize XL: , " + bstrCachePath, logFilePath);
                    return;
                }
                if (tempXL.Length == 1)
                {
                    if (File.Exists(tempXL[0]) == true)
                    {
                        //try
                        //{
                        //    File.Delete(XLTemplateFile);
                        //}
                        //catch (Exception ex)
                        //{
                        //    Utility.Log("Delete Failed, " + XLTemplateFile + ":" + ex.Message, logFilePath);
                        //    return;
                        //}
                        try
                        {
                            System.IO.File.Move(tempXL[0], XLTemplateFile);
                        }
                        catch (Exception ex)
                        {
                            Utility.Log("Move Failed, " + XLTemplateFile + ":" + ex.Message, logFilePath);
                            return;
                        }
                    }
                    //XLTemplateFile = XlFiles[0];
                }
                else
                {
                    Utility.Log("More than 1 XL found inside the cache, cannot sanitize XL: , " + bstrCachePath, logFilePath);
                    return;

                }
            }
            else
            {
                Utility.Log("XLTemplateFile is already downloaded and found.., " + XLTemplateFile, logFilePath);
            }
            
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            if (File.Exists(XLTemplateFile) == false)
            {
                Utility.Log("Could not find the XL template in Cache.." + XLTemplateFile, logFilePath);
                return;
            }

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
                //Clean up objects
                xlApp.Quit();
                xlApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        // Rename the excel sheet names using the based_on property.
        public static void RenameExcelSheet(Microsoft.Office.Interop.Excel._Application xlApp, string sFileName, String logFilePath)
        {   
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

                        Utlity.Log(sheet.Name, logFilePath);
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
                        if (basedOnDictionaryBOAKey.ContainsKey(name) == true)
                        {
                            String value = "";
                            basedOnDictionaryBOAKey.TryGetValue(name, out value);
                            //Rename the sheet 
                            excelWorkSheet.Name = value + "." + extn; ;
                        }

                        Marshal.ReleaseComObject(sheet);
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
                Utility.Log("Export Excel Failed: " + ex.Message,logFilePath);
            }            
        }

        public static bool SanitizeSheetInXL(Microsoft.Office.Interop.Excel._Application xlApp,
            String outputXLfileName,
          String logFilePath)
        {
            //Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlApp.Visible = false;
            //xlApp.DisplayAlerts = false;

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
                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true || 
                    sheet.Name.Equals("MASTER ASSEMBLY",StringComparison.OrdinalIgnoreCase) == true)
                {
                    try
                    {
                            Utlity.Log("sheet.Name" + ":::" + sheet.Name, logFilePath);
                        foreach(string key in basedOnDictionaryBOAKey.Keys) {
                            //WriteFeatureSheet(xlWorkbook, sheet, sheet.Name, logFilePath);
                            String value = "";
                             basedOnDictionaryBOAKey.TryGetValue(key,out value);
                             FindAndReplace(xlWorkbook, sheet, key, value, logFilePath);
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
                }else
                {
                    WriteSheet(xlWorkbook, sheet, logFilePath);

                    try
                    {
                        //Utlity.Log(sheet.Name, logFilePath);
                        if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                        {
                            if (partEnablementDictionary.ContainsKey(sheet.Name) == false)
                            {
                                Utlity.Log("SanitizeSheetInXL: partEnablementDictionary Key " + sheet.Name, logFilePath);
                                Utlity.Log("SanitizeSheetInXL: partEnablementDictionary Value " + "HIDDEN", logFilePath);
                                partEnablementDictionary.Add(sheet.Name, false);
                            }
                        }
                        else
                        {
                            if (partEnablementDictionary.ContainsKey(sheet.Name) == false)
                            {
                                Utlity.Log("SanitizeSheetInXL: partEnablementDictionary Key " + sheet.Name, logFilePath);
                                Utlity.Log("SanitizeSheetInXL: partEnablementDictionary Value " + "VISIBLE", logFilePath);
                                partEnablementDictionary.Add(sheet.Name, true);
                            }
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

            ////quit and release
            //if (xlApp != null) xlApp.DisplayAlerts = true;
            //if (xlApp != null) xlApp.Quit();
            //Marshal.ReleaseComObject(xlApp);
            //xlApp = null;

            return true;
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
                        Utlity.Log("oldName is Empty...", logFilePath);
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

                    if (basedOnDictionaryBOAKey.ContainsKey(name) == true)
                    {
                        String value = "";
                        basedOnDictionaryBOAKey.TryGetValue(name, out value);
                        xlRange.Cells[i, 9].Value2 = value + "." + extn;
                    }

                }
            }
            catch (Exception ex)
            {
                Utlity.Log("Exception in Writing Sheet..." + ex.Message, logFilePath);
                Utlity.Log("Exception in Writing Sheet..." + ex.StackTrace, logFilePath);                
            }

            xlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

            //Marshal.ReleaseComObject(xlWorkSheet);
            //xlWorkSheet = null;

        }


        private static void WriteFeatureSheet(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel.Worksheet sheet, String ocurrenceName, String logFilePath)
        {
            if (sheet == null)
            {
                Utlity.Log("WriteFeatureSheet: sheet is Empty" + ocurrenceName, logFilePath);
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
                if (xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1] != null)
                {
                    try
                    {
                        String partName = xlRange.Cells[i, 1].Value2;
                        if (partName != null && partName.Equals("") == false)
                        {
                            String namewithoutExtn = partName.Split('.')[0];
                            if (basedOnDictionaryBOAKey.ContainsKey(namewithoutExtn) == false)
                            {
                                Utlity.Log("could not find part_name in basedOnDictionary " + namewithoutExtn, logFilePath);
                                continue;
                            }
                            String based_on_value="";
                            var element = basedOnDictionaryBOAKey.TryGetValue(namewithoutExtn, out based_on_value);

                            if (element != null)
                            {
                                // SET VALUE
                                xlRange.Cells[i, 1].Value2 = based_on_value;                                
                            }
                            else
                            {
                                Utlity.Log("WriteFeatureSheet: " + "Could Not Find " + partName, logFilePath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("partName" + ex.Message, logFilePath);
                    }
                }

            }
            xlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;

            //Marshal.ReleaseComObject(sheet);
            //sheet = null;


        }

        public static void FindAndReplace(Microsoft.Office.Interop.Excel.Workbook xlWorkbook, Microsoft.Office.Interop.Excel.Worksheet sheet,
            String find, String replace, String logFilePath)
        {

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
                true, Type.Missing, Type.Missing, Type.Missing);

            xlWorkbook.Save();

            Marshal.ReleaseComObject(xlRange);
            xlRange = null;



        }

        // another copy from SolidEdgeData1 class
        public static void traverseAssembly(String assemblyFileName, String logFilePath)
        {

            // 19/8 - Purpose - To find repetitive BomLines & Not Add into bomLineList (Object Store)
            // If Added Again into Object Store, Issues Arise in TreeView in the UI.            

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
            String AbsolutePath = document.AbsolutePath;
            objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
            Utlity.Log("bstrItemId: " + bstrItemId, logFilePath);
            Utlity.Log("bstrItemRev: " + bstrItemRev, logFilePath);
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

                AbsolutePath = linkDocument.AbsolutePath;
                objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
                Utlity.Log("bstrItemId: " + bstrItemId, logFilePath);
                Utlity.Log("bstrItemRev: " + bstrItemRev, logFilePath);
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

                String AbsolutePath = linkDocument.AbsolutePath;

                objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
                Utlity.Log("bstrItemId: " + bstrItemId, logFilePath);
                Utlity.Log("bstrItemRev: " + bstrItemRev, logFilePath);
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


    }
}
