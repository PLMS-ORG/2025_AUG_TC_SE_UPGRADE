using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using AddToTc.CTD;
using Creo_TC_Live_Integration.TcDataManagement;
using DemoAddInTC.se;
using DemoAddInTC.utils;
using Log;
using Microsoft.Office.Interop.Excel;
using Teamcenter.Soa.Client;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;

namespace DemoAddInTC.controller
{
    class SanitizeXL_PostClone
    {
        public static Dictionary<String, String> basedOnDictionaryBOAKey = new Dictionary<String, String>();
        public static Dictionary<String, String> basedOnDictionaryitemIDKey = new Dictionary<String, String>();
        public static List<String> itemIDCollection = new List<string>();
        public static SolidEdgeFramework.SolidEdgeTCE objSEEC = null;
        public static Dictionary<String, bool> partEnablementDictionary = new Dictionary<string, bool>();

        public static void read_all_items_in_cache(String assemblyFileName, String username, string pass, string grp,string stageDir)
        {
            basedOnDictionaryBOAKey.Clear();
            basedOnDictionaryitemIDKey.Clear();
            itemIDCollection.Clear();

            //START SEEC session here using SEECAdaptor
            //SEECAdaptor.SetCredentials(username, pass, grp);
            //SEECAdaptor.LoginToTeamcenter();
            
            objSEEC = SEECAdaptor.getSEECObject();
            if (objSEEC == null)
            {
                log.write(logType.INFO, "objSEEC is null.. " + assemblyFileName);
                return;

            }
                    
            String bstrCachePath = SEECAdaptor.GetPDMCachePath();

            if (bstrCachePath == "" || bstrCachePath == null)
            {
                log.write(logType.INFO, "bstrCachePath is null.. " + assemblyFileName);
                return;
            }

            log.write(logType.INFO,"read_all_items_in_cache: Traversing the assembly: " + assemblyFileName);
            traverseAssembly(assemblyFileName);

            write_to_InputToChangeOwnership(itemIDCollection, stageDir);
            
            log.write(logType.INFO, "read_all_items_in_cache: Collecting based on Property: " + assemblyFileName);
            collect_based_on_property(itemIDCollection);

            log.write(logType.INFO, "read_all_items_in_cache: read and Sanitize the XL: " + assemblyFileName); // -- 
            read_and_sanitize_XL(bstrCachePath,assemblyFileName);

            String xlFile = Path.ChangeExtension(assemblyFileName, ".xlsx");
            if (File.Exists(xlFile) == true)
            {
                log.write(logType.INFO, "read_all_items_in_cache: Upload the excel back to Teamcenter: " + assemblyFileName);
                TCUtils.uploadExcelToTC(username, pass, grp,"", xlFile);

                String Revision = Utlity.getPropertyValue("ITEM_REVISION_TYPE"); //"P7MCPgPartRevision"
                log.write(logType.INFO, "ITEM_REVISION_TYPE: " + Revision);
                if (Revision.Equals("ltc4_ItemRevision", StringComparison.OrdinalIgnoreCase) == true)
                {
                    log.write(logType.INFO, "updateSuffixProperty ");
                    updateSuffixProperty(itemIDCollection);
                }
            }
            else
            {
                log.write(logType.INFO, "read_all_items_in_cache: " + "no xl file is found in cache");
            }

            // Kill SEEC session here using SEECAdaptor
            SEECAdaptor.KillSESession();
        }


        private static void write_to_InputToChangeOwnership(List<string> itemIDCollection, String stageDir)
        {
            if (itemIDCollection != null && itemIDCollection.Count > 0)
            {

                String InputToChangeOwnership = Path.Combine(stageDir, "InputToChangeOwnership.txt");
                log.write(logType.INFO, "Writing to InputToChangeOwnership Text File..");
                using (StreamWriter file = new StreamWriter(InputToChangeOwnership))
                    foreach (var entry in itemIDCollection)
                    {
                        String item_ID = entry.Split('~')[0];
                        String rev_ID = entry.Split('~')[1];

                        file.WriteLine("{0}~{1}", item_ID, rev_ID);
                    }
            }
        }

        // update the part var type property
        public static void updateSuffixProperty(List<String> itemIDCollectio)
        {
            if (itemIDCollection == null && itemIDCollection.Count == 0)
            {
                log.write(logType.INFO, "updateSuffixProperty: itemIDCollection is Empty:");
                return;
            }

            if (partEnablementDictionary == null && partEnablementDictionary.Count == 0)
            {
                log.write(logType.ERROR, "updateSuffixProperty: partEnablementDictionary is Empty:");
                return;
            }
            log.write(logType.INFO, "updateSuffixProperty: partEnablementDictionary Count :" + partEnablementDictionary.Count);
            foreach (String key in partEnablementDictionary.Keys)
            {
                // key is sheet name - which is itemid.extn
                bool value;
                partEnablementDictionary.TryGetValue(key, out value);
                log.write(logType.INFO, "updateSuffixProperty: partEnablementDictionary Key :" + key);
                log.write(logType.INFO, "updateSuffixProperty: partEnablementDictionary Value :" + value);
                if (objSEEC == null)
                {
                    log.write(logType.ERROR, "updateSuffixProperty: objSEEC is NULL:");
                    return;
                }

                String bStrCachePath = "";
                objSEEC.GetPDMCachePath(out bStrCachePath);

                String documentFileName = Path.Combine(bStrCachePath, key);
                String bstrItemId = "";
                String bstrItemRev = "";

                objSEEC.GetDocumentUID(documentFileName, out bstrItemId, out bstrItemRev);

                log.write(logType.INFO, "updateSuffixProperty::documentFileName: " + documentFileName);
                log.write(logType.INFO, "updateSuffixProperty::bstrItemId: " + bstrItemId);
                log.write(logType.INFO, "updateSuffixProperty::bstrItemRev: " + bstrItemRev);

                if (bstrItemId == null || bstrItemId.Equals("") == true || bstrItemRev == null ||
                    bstrItemRev.Equals("") == true)
                {
                    log.write(logType.INFO, "bstrItemId is empty: " + bstrItemId);
                    log.write(logType.INFO, "bstrItemRev is empty: " + bstrItemRev);
                    continue;
                }

                ModelObject itemRevMO = DownloadDatasetNamedReference.getItemRevisionQuery(bstrItemId, bstrItemRev);

                if (itemRevMO == null)
                {

                    log.write(logType.INFO, "Item Rev Model Object is NULL: " + bstrItemId);
                    continue;
                }

                try
                {
                    if (value == true)
                    {
                        log.write(logType.INFO, "ltc4_Part_Var_Type Property Value:" + "Variable");
                        TCUtils.setIRProperty(itemRevMO, "Variable", "ltc4_Part_Var_Type");
                        try
                        {
                            // 14-03-2024: Suffix is hardcoded as "CTD" when we run CTD in Dispatcher Mode
                            log.write(logType.INFO,"Suffix property from CTD : " + "CTD");
                            TCUtils.setIRProperty(itemRevMO, "CTD", "ltc4_Suffix1");
                        }
                        catch (Exception)
                        {
                            log.write(logType.ERROR,"Error encountered while updating suffix");
                        }
                        
                    }
                    else
                    {
                        log.write(logType.INFO, "ltc4_Part_Var_Type Property Value:" + "constant");
                        TCUtils.setIRProperty(itemRevMO, "constant", "ltc4_Part_Var_Type");
                    }

                }

                catch (Exception e)
                {
                    log.write(logType.INFO, "exception in getAllDataSet function" + e.Message);
                }

            }



        }

        public static void collect_based_on_property(List<String>itemIDCollection)       
        {
            foreach (String item in itemIDCollection)
            {
                String item_ID = item.Split('~')[0];
                String rev_ID = item.Split('~')[1];

               log.write(logType.INFO,"collect_based_on_property: getItemRevisionQuery:" + item_ID + "~" + rev_ID);
               ModelObject itemRevMO = DownloadDatasetNamedReference.getItemRevisionQuery(item_ID, rev_ID);

                if (itemRevMO == null)
                {

                    log.write(logType.ERROR, "Item Rev Model Object is NULL: " + item_ID);
                    return;
                }

                try
                {

                    ItemRevision itemRev = (ItemRevision)itemRevMO;
                    if (itemRev == null)
                    {
                        log.write(logType.ERROR, "collect_based_on_property: " + "itemRev is null...");
                        return;
                    }

                    String based_on = itemRev.Based_on;
                    String based_on_actual = based_on.Split('/')[0];

                    if (basedOnDictionaryBOAKey.ContainsKey(based_on_actual) == false)
                    {
                        log.write(logType.INFO, "basedOnDictionaryBOAKey: " + "based_on_actual..." + based_on_actual);
                        log.write(logType.INFO, "basedOnDictionaryBOAKey: " + "item_ID..." + item_ID);
                        basedOnDictionaryBOAKey.Add(based_on_actual, item_ID);
                    }

                    if (basedOnDictionaryitemIDKey.ContainsKey(item_ID) == false)
                    {
                        log.write(logType.INFO, "basedOnDictionaryitemIDKey: " + "based_on_actual..." + based_on_actual);
                        log.write(logType.INFO, "basedOnDictionaryitemIDKey: " + "item_ID..." + item_ID);
                        basedOnDictionaryitemIDKey.Add(item_ID, based_on_actual);
                    }               
                }
                catch (Exception e)
                {
                    log.write(logType.ERROR, "exception in basedOnDictionaryBOAKey function" + e.Message);
                }
            }
        }

        public static void read_and_sanitize_XL(String bstrCachePath, String assemblyFileName)
        {
            String XLTemplateFile = System.IO.Path.ChangeExtension(assemblyFileName, ".xlsx");
            String bstrItemId = "";
            String bstrItemRev = "";

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
                    log.write(logType.INFO, "Deleting xLFile: , " + xLFile);
                    File.Delete(xLFile);                   
                }

                String[] tempXL = Directory.GetFiles(bstrCachePath, "*", SearchOption.AllDirectories)
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
                                             .ToArray();

                if (tempXL.Length == 0)
                {
                    log.write(logType.ERROR, "No XL found inside the cache, cannot sanitize XL: , " + bstrCachePath);
                    return;
                }
                if (tempXL.Length == 1)
                {
                    if (File.Exists(tempXL[0]) == true)
                    {
                        try
                        {
                            System.IO.File.Move(tempXL[0], XLTemplateFile);
                        }
                        catch (Exception ex)
                        {
                            log.write(logType.ERROR, "Move Failed, " + XLTemplateFile + ":" + ex.Message);
                            return;
                        }
                    }
                    //XLTemplateFile = XlFiles[0];
                }
                else
                {
                    log.write(logType.ERROR, "More than 1 XL found inside the cache, cannot sanitize XL: , " + bstrCachePath);
                    return;

                }
            }
            else
            {
                log.write(logType.INFO, "XLTemplateFile is already downloaded and found.., " + XLTemplateFile);
            }

            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            if (File.Exists(XLTemplateFile) == false)
            {
                log.write(logType.INFO, "Could not find the XL template in Cache.." + XLTemplateFile);
                return;
            }

            try
            {
                // change the sheet names
                RenameExcelSheet(xlApp, XLTemplateFile);
            }
            catch (Exception ex)
            {
                log.write(logType.INFO, "Could not rename sheets in Excel File.." + XLTemplateFile);
            }

            try
            {
                // open the FEATURES sheet and modify the partname column.
                // open the MASTER ASSEMBLY and modify -- FullName , AbsolutePath, DocNum
                SanitizeSheetInXL(xlApp, XLTemplateFile);
            }
            catch (Exception e)
            {
                log.write(logType.INFO, "Could not SanitizeSheetInXL.." + XLTemplateFile);
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
        public static void RenameExcelSheet(Microsoft.Office.Interop.Excel._Application xlApp, string sFileName)
        {   
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
            Microsoft.Office.Interop.Excel._Worksheet excelWorkSheet;

            try
            {
                excelWorkbook = xlApp.Workbooks.Open(sFileName);
                if (excelWorkbook.Sheets.Count > 0)
                {
                    String ltcCustomSheetName = Utlity.getLTCCustomSheetName();
                    foreach (Microsoft.Office.Interop.Excel._Worksheet sheet in excelWorkbook.Sheets)
                    {
                        // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                        if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                        {
                            if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                            {
                                log.write(logType.INFO,"Skipping Custom LTC Sheet: " + sheet.Name);
                                Marshal.ReleaseComObject(sheet);
                                continue;
                            }
                        }

                        log.write(logType.INFO, sheet.Name);
                        if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true ||
                            sheet.Name.Equals("MASTER ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            log.write(logType.INFO, "Skipping : " + sheet.Name);
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
                log.write(logType.ERROR,"Export Excel Failed: " + ex.Message);
            }            
        }

        public static bool SanitizeSheetInXL(Microsoft.Office.Interop.Excel._Application xlApp,
                                             String outputXLfileName)  
        {
            //Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlApp.Visible = false;
            //xlApp.DisplayAlerts = false;

            log.write(logType.INFO, "Inside SanitizeSheetInXL");
                       
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            FileInfo f = new FileInfo(outputXLfileName);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            if (f.Exists == true)
            {
                log.write(logType.INFO, "File Already Exists,");
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
                log.write(logType.ERROR, "xlWorkBook is NULL");
                return false;
            }
            // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
            String ltcCustomSheetName = utils.Utlity.getLTCCustomSheetName();
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in xlWorkbook.Worksheets)
            {
                // 29 - SEPT - Ignore Custom Sheet Names Added By LTC Designers.
                if (ltcCustomSheetName != null && ltcCustomSheetName.Equals("") == false)
                {
                    if (sheet.Name.StartsWith(ltcCustomSheetName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        log.write(logType.INFO, "Skipping Custom LTC Sheet: " + sheet.Name);
                        Marshal.ReleaseComObject(sheet);
                        continue;
                    }
                }

                log.write(logType.INFO, sheet.Name);
                if (sheet.Name.Equals("FEATURES", StringComparison.OrdinalIgnoreCase) == true || 
                    sheet.Name.Equals("MASTER ASSEMBLY",StringComparison.OrdinalIgnoreCase) == true)
                {
                    try
                    {
                        log.write(logType.INFO, "sheet.Name" + ":::" + sheet.Name);
                        foreach(string key in basedOnDictionaryBOAKey.Keys) {
                            //WriteFeatureSheet(xlWorkbook, sheet, sheet.Name, logFilePath);
                            String value = "";
                             basedOnDictionaryBOAKey.TryGetValue(key,out value);
                             FindAndReplace(xlWorkbook, sheet, key, value);
                        }
               
                    }
                    catch (Exception ex)
                    {
                        //Marshal.ReleaseComObject(xlWorkbook);
                        //xlWorkbook = null;
                        //Marshal.ReleaseComObject(workbooks);
                        //workbooks = null;
                        Marshal.ReleaseComObject(sheet);
                        log.write(logType.INFO, "SanitizeSheetInXL: Exception " + ex.Message);
                        continue;
                    }
                }else
                {
                    WriteSheet(xlWorkbook, sheet);

                    try
                    {
                        //Utlity.Log(sheet.Name, logFilePath);
                        if (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                        {
                            if (partEnablementDictionary.ContainsKey(sheet.Name) == false)
                            {
                                log.write(logType.INFO, "SanitizeSheetInXL: partEnablementDictionary Key " + sheet.Name);
                                log.write(logType.INFO, "SanitizeSheetInXL: partEnablementDictionary Value " + "HIDDEN");
                                partEnablementDictionary.Add(sheet.Name, false);
                            }
                        }
                        else
                        {
                            if (partEnablementDictionary.ContainsKey(sheet.Name) == false)
                            {
                                log.write(logType.INFO, "SanitizeSheetInXL: partEnablementDictionary Key " + sheet.Name);
                                log.write(logType.INFO, "SanitizeSheetInXL: partEnablementDictionary Value " + "VISIBLE");
                                partEnablementDictionary.Add(sheet.Name, true);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.write(logType.ERROR, ex.Message);
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
            Microsoft.Office.Interop.Excel.Worksheet Sheet)
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
                        log.write(logType.INFO, "oldName is Empty...");
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
                log.write(logType.ERROR, "Exception in Writing Sheet..." + ex.Message);
                log.write(logType.ERROR, "Exception in Writing Sheet..." + ex.StackTrace);                
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

                            if (element == true)
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
            String find, String replace)
        {

            if (sheet == null)
            {
                log.write(logType.INFO, "FindAndReplace: sheet is Empty");
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
        public static void traverseAssembly(String assemblyFileName)
        {

            // 19/8 - Purpose - To find repetitive BomLines & Not Add into bomLineList (Object Store)
            // If Added Again into Object Store, Issues Arise in TreeView in the UI.            

            SEECAdaptor.InitializeSolidEdgeRevisionManagerSession();
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SEECAdaptor.objReviseApp;
            SolidEdge.RevisionManager.Interop.Document document = null;
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;
            SolidEdge.RevisionManager.Interop.Document linkDocument = null;
            try
            {
                document = objReviseApp.Open(assemblyFileName);
                if (document == null)
                {
                    log.write(logType.ERROR, "traverseAssembly: " + "Document is NULL");
                    return;
                }
            }
            catch (Exception ex)
            {
                log.write(logType.ERROR, "traverseAssembly: " + ex.Message);
                return;
            }

            String bstrItemId = "";
            String bstrItemRev = "";
            String AbsolutePath = document.AbsolutePath;

            if (objSEEC == null)
            {
                log.write(logType.INFO, "objSEEC is NULL : ");
                return;
            }
            if(File.Exists(AbsolutePath) == false)
            {
                log.write(logType.INFO, "Absolute Path is not Exists: " + AbsolutePath);
                return;
            }
            objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
            log.write(logType.INFO, "bstrItemId: " + bstrItemId);
            log.write(logType.INFO, "bstrItemRev: " + bstrItemRev);
            if (itemIDCollection.Contains(bstrItemId + "~" + bstrItemRev) == false)
            {
                itemIDCollection.Add(bstrItemId + "~" + bstrItemRev);
            }



            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                log.write(logType.INFO, "No Linked Documents in : " + assemblyFileName);
                return;

            }

            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];
                if (linkDocument.FullName.EndsWith(".xlsx") == true)
                {
                    log.write(logType.INFO, "Skipping: " + linkDocument.FullName);
                    continue;
                }

                log.write(logType.INFO, "FullName: " + linkDocument.FullName);
                //Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                //Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                //Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                //Utlity.Log("Status: " + linkDocument.Status, logFilePath);

                AbsolutePath = linkDocument.AbsolutePath;
                objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
                log.write(logType.INFO, "bstrItemId: " + bstrItemId);
                log.write(logType.INFO, "bstrItemRev: " + bstrItemRev);
                if (itemIDCollection.Contains(bstrItemId + "~" + bstrItemRev) == false)
                {
                    itemIDCollection.Add(bstrItemId + "~" + bstrItemRev);
                }


                //bomLineList.Add(bl);
                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument);

                }
            }

            SEECAdaptor.killRevisionManager();

        }

        private static void traverseLinkDocuments(SolidEdge.RevisionManager.Interop.Document linkDocument)
        {
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                log.write(logType.ERROR, "No Linked Documents in : " + linkDocument.FullName);
                return;
            }

            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                log.write(logType.INFO, "FullName: " + linkDocument.FullName);
                //Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                //Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                //Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                //Utlity.Log("Status: " + linkDocument.Status, logFilePath);

                String bstrItemId = "";
                String bstrItemRev = "";

                String AbsolutePath = linkDocument.AbsolutePath;

                objSEEC.GetDocumentUID(AbsolutePath, out bstrItemId, out bstrItemRev);
                log.write(logType.INFO, "bstrItemId: " + bstrItemId);
                log.write(logType.INFO, "bstrItemRev: " + bstrItemRev);
                if (itemIDCollection.Contains(bstrItemId + "~" + bstrItemRev) == false)
                {
                    itemIDCollection.Add(bstrItemId + "~" + bstrItemRev);
                }

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument);
                }
            }


        }


    }
}
