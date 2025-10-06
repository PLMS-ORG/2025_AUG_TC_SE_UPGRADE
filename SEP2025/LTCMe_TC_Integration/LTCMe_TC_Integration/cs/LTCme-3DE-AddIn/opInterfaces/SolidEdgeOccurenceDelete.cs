using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DemoAddInTC.opInterfaces
{
    // 02/SEPT - ADDED based on Simone Mail.
    // Reads all Disabled Parts in Assembly -- Based on Sheet HIDE status
    // Traverses the Assembly and Inactivates the Occurence in SE if its Hidden by ADMIN
    // Reference Material - https://community.plm.automation.siemens.com/t5/Solid-Edge-Developer-Forum/How-to-Suppress-a-part-in-Assembly-just-like-Solidworks-does/m-p/312838#M7034

    class SolidEdgeOccurenceDelete
    {

        // CAUTION - assemblyFileName - OLD ASSEMBLY PATH
        public static void process(String publishedFolder, String assemblyFilePath, String logFilePath)
        {

            //String[] excelFiles = Directory.GetFiles(publishedFolder, "*", SearchOption.AllDirectories)
            //                            .Select(path => Path.GetFullPath(path))
            //                            .Where(x => (x.EndsWith(".xlsx")))
            //                            .ToArray();
            //if (excelFiles == null || excelFiles.Length == 0)
            //{
            //    Utlity.Log("No Excel Files Found In " + publishedFolder, logFilePath);
            //    return ;
            //}
            String assemblyFile = Path.GetFileName(assemblyFilePath);
            String assemblyFileInPublishedFolder = Path.Combine(publishedFolder, assemblyFile);
            if (System.IO.File.Exists(assemblyFileInPublishedFolder) == false)
            {
                Utlity.Log("FILE NOT FOUND " + assemblyFileInPublishedFolder, logFilePath);
                return;
            }

            String xlFilePath = Path.ChangeExtension(assemblyFileInPublishedFolder, ".xlsx");

            if (System.IO.File.Exists(xlFilePath) == false)
            {
                Utlity.Log("FILE NOT FOUND " + xlFilePath, logFilePath);
                return;
            }

            //String xlFilePath = excelFiles[0];

            Dictionary<String, bool> OccurenceEnablementDictionary = readOccurenceEnablementInfoFromTemplateExcel(xlFilePath, logFilePath);

            //String assemblyFileName = Path.ChangeExtension(xlFilePath, ".asm");
            readOccurences(assemblyFileInPublishedFolder, logFilePath, OccurenceEnablementDictionary);

        }

        public static Dictionary<String, bool> readOccurenceEnablementInfoFromTemplateExcel(String xlFilePath, String logFilePath)
        {
            Dictionary<String, bool> OccurenceEnablementDictionary = new Dictionary<string, bool>();


            //Utlity.Log("----------------------------------------------------------", logFilePath);
            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            FileInfo f = new FileInfo(xlFilePath);
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            if (f.Exists == true)
            {
                workbooks = xlApp.Workbooks;
                Utlity.Log("readOccurenceEnablementInfoFromTemplateExcel", logFilePath);
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
            String topLineAssembly = Path.ChangeExtension(xlFilePath, ".asm");
            String topLine = Path.GetFileName(topLineAssembly);
            //Utlity.Log("WorkSheet Count: " + xlWorkbook.Worksheets.Count.ToString(), logFilePath);
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkbook.Worksheets;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
            {

               

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
                        OccurenceEnablementDictionary.Add(sheet.Name, false);
                    }
                    else
                    {
                        OccurenceEnablementDictionary.Add(sheet.Name, true);
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }
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

            if (xlApp != null) xlApp.DisplayAlerts = true;
            //Utlity.Log("----------------------------------------------------------", logFilePath);
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

            return OccurenceEnablementDictionary;
        }


        public static void readOccurences(String assemblyFileName, String logFilePath, Dictionary<String, bool> OccurenceEnablementDictionary)
        {
            if (assemblyFileName == null || assemblyFileName.Equals("") == true)
            {
                Utlity.Log("SolidEdgeOccurenceDelete: No Assembly File in the Published Folder: ", logFilePath);
                return;
            }

            if (File.Exists(assemblyFileName) == false)
            {
                Utlity.Log("SolidEdgeOccurenceDelete: No Assembly File in the Published Folder: " + assemblyFileName, logFilePath);
                return;
            }

            if (OccurenceEnablementDictionary == null || OccurenceEnablementDictionary.Count == 0)
            {
                Utlity.Log("SolidEdgeOccurenceDelete: traverseAssembly: " + "OccurenceEnablementDictionary is Empty", logFilePath);
                return;
            }

            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
            {
                Utlity.Log("DEBUG - objApp is NULL : ", logFilePath);
                return;
            }

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;


            


            try
            {
                objDocuments = objApp.Documents;
                objApp.DisplayAlerts = false;
                //objApp.Visible = false;
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                //objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;
                if (objAssemblyDocument == null)
                {
                    Utlity.ResetAlerts(objApp, true, logFilePath);
                    Utlity.Log("DEBUG - InputFile is NOT Opened : ", logFilePath);
                    return;
                }


                if (objAssemblyDocument != null)
                {
                    // This is for Top Assembly Alone
                    // Utlity.Log("AssemDoc.Name : " + objAssemblyDocument.Name, logFilePath);                    
                    occurrences = objAssemblyDocument.Occurrences;
                    
                    if (occurrences == null || occurrences.Count == 0)
                    {
                        Utlity.ResetAlerts(objApp, true, logFilePath);
                        Utlity.Log("occurrences is Empty--", logFilePath);
                        return;
                    }

                    for (int i = 1; i <= occurrences.Count; i++)
                    {
                        occurrence = occurrences.Item(i);
                        String occurenceName = occurrence.Name;
                        String[] occArr = occurenceName.Split(':');
                        if (occArr.Length == 2)
                        {
                            occurenceName = occArr[0];
                        }
                        else
                        {
                            Utlity.Log("Cant Parse OccurenceName--", logFilePath);
                            continue;
                        }
                        

                        //Utlity.Log("-----------------------------------------", logFilePath);
                        //Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        //Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        //Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            
                            if (OccurenceEnablementDictionary.ContainsKey(occurenceName) == true)
                            {
                                bool value = false;
                                bool Success = OccurenceEnablementDictionary.TryGetValue(occurenceName, out value);

                                if (Success == true)
                                {                                    
                                    OccurrenceActivate(occurrence, value, logFilePath);                                 
                                }

                            }

                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            if (OccurenceEnablementDictionary.ContainsKey(occurenceName) == true)
                            {                                
                                
                                bool value = false;
                                bool Success = OccurenceEnablementDictionary.TryGetValue(occurenceName, out value);

                                if (Success == true)
                                {
                                    OccurrenceActivate(occurrence, value, logFilePath);
                                }

                            }
                        }
                        else if (occurrence.OccurrenceFileName.EndsWith(".asm") == true)
                        {
                            SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;
                            assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            bool value = false;
                            if (OccurenceEnablementDictionary.ContainsKey(occurenceName) == true)
                            {

                                
                                bool Success = OccurenceEnablementDictionary.TryGetValue(occurenceName, out value);

                                if (Success == true)
                                {
                                    OccurrenceActivate(occurrence, value, logFilePath);
                                }

                            }
                            //if (value == true)
                            {
                                // Traverse the assembly Only when the parent ASM is Visible.Else no need to Traverse Inside.
                                traverseAssembly(assemDoc1, logFilePath, OccurenceEnablementDictionary);
                            }
                        }

                        //Utlity.Log("-----------------------------------------", logFilePath);
                    }
                }
                else
                {
                    Utlity.ResetAlerts(objApp, true, logFilePath);
                    return;
                }

            }
            catch (Exception ex)
            {
                Utlity.ResetAlerts(objApp, true, logFilePath);
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.Source, logFilePath);
            }

            SaveAndCloseAssembly(objAssemblyDocument, logFilePath);
            if (objApp != null) objApp.DisplayAlerts = true;  
        }

        private static void OccurrenceActivate(SolidEdgeAssembly.Occurrence occurrence, bool value, String logFilePath)
        {
            try
            {
                Utlity.Log(occurrence.Name + " : Visiblility : " + value, logFilePath);                
                occurrence.Visible = value;
                occurrence.IncludeInBom = value;
                occurrence.IncludeInInterference = value;
                occurrence.IncludeInPhysicalProperties = value;
                occurrence.DisplayInDrawings = value;
                occurrence.DisplayInSubAssembly = value;
                occurrence.Activate = value;
                
                
            }
            catch (Exception ex)
            {
                Utlity.Log("OccurrenceActivate- Exception: " + ex.Message, logFilePath);
            }
        }


        private static void traverseAssembly(SolidEdgeAssembly.AssemblyDocument assemDoc, String logFilePath, Dictionary<String, bool> OccurenceEnablementDictionary)
        {
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            if (assemDoc == null)
            {
                Utlity.Log("assemDoc is Empty: " + assemDoc.Name, logFilePath);
                return;
            }
            //Utlity.Log("assemDoc.Name: " + assemDoc.Name, logFilePath);            

            occurrences = assemDoc.Occurrences;
            //Utlity.Log("occurrences.Count: " + occurrences.Count, logFilePath);
            for (int i = 1; i <= occurrences.Count; i++)
            {
                occurrence = occurrences.Item(i);
                String occurenceName = occurrence.Name;
                String[] occArr = occurenceName.Split(':');
                if (occArr.Length == 2)
                {
                    occurenceName = occArr[0];
                }
                else
                {
                    Utlity.Log("Cant Parse OccurenceName--", logFilePath);
                    continue;
                }
               

                //Utlity.Log("-----------------------------------------", logFilePath);
                Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                int occurenceQty = occurrence.Quantity;
                //Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                String ocurenceFileName = occurrence.OccurrenceFileName;
                //Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    if (OccurenceEnablementDictionary.ContainsKey(occurenceName) == true)
                    {
                        bool value = false;
                        bool Success = OccurenceEnablementDictionary.TryGetValue(occurenceName, out value);

                        if (Success == true)
                        {
                            OccurrenceActivate(occurrence, value, logFilePath);
                        }

                    }
                }
                else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    if (OccurenceEnablementDictionary.ContainsKey(occurenceName) == true)
                    {
                        bool value = false;
                        bool Success = OccurenceEnablementDictionary.TryGetValue(occurenceName, out value);

                        if (Success == true)
                        {
                            OccurrenceActivate(occurrence, value, logFilePath);
                        }

                    }
                }
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    SolidEdgeAssembly.AssemblyDocument assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;

                    traverseAssembly(assemDoc1, logFilePath, OccurenceEnablementDictionary);


                }
            }

        }


        private static void SaveAndCloseAssembly(SolidEdgeAssembly.AssemblyDocument assemblyDoc, String logFilePath)
        {
            try
            {
                if (assemblyDoc.ReadOnly == false)
                {
                    assemblyDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }

            try
            {
                //if (assemblyDoc.ReadOnly == false)
                //{
                //    assemblyDoc.Close(true);
                //}
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }

        }


        private static void savePart(SolidEdgePart.PartDocument partDoc, String logFilePath)
        {
            try
            {
                if (partDoc.ReadOnly == false)
                {
                    partDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }
        }

        private static void saveSheet(SolidEdgePart.SheetMetalDocument sheetMetalDoc, String logFilePath)
        {
            try
            {
                if (sheetMetalDoc.ReadOnly == false)
                {
                    sheetMetalDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }
        }
        
    }
}
