using _3DAutomaticUpdate.controller;
using SolidEdgeCommunity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate.opInterfaces
{
    class SolidEdgeOccurenceDelete_1 : IsolatedTaskProxy
    {
        List<String> occurenceList = new List<string>();

        public void SolidEdgeOccurrenceDeleteFromExcelSTAT(String assemblyFileName, String logFilePath)
        {
            InvokeSTAThread<String, String>(SolidEdgeOccurrenceDeleteFromExcel, assemblyFileName, logFilePath);
        }

        private void OccurrenceDelete(SolidEdgeAssembly.Occurrence occurrence, bool value, String logFilePath)
        {
            try
            {
                Utility.Log(occurrence.Name + " : Deleting..", logFilePath);
                occurrence.Visible = value;
                occurrence.IncludeInBom = value;
                occurrence.IncludeInInterference = value;
                occurrence.IncludeInPhysicalProperties = value;
                occurrence.DisplayInDrawings = value;
                occurrence.DisplayInSubAssembly = value;

                occurrence.Activate = value;
                occurrence.Delete();

            }
            catch (Exception ex)
            {
                Utility.Log("OccurrenceActivate- Exception: " + ex.Message, logFilePath);
            }
        }



        [STAThread]
        public void SolidEdgeOccurrenceDeleteFromExcel(String assemblyFileName, String logFilePath)
        {

            ReadAllOccurrencesFromAssembly(assemblyFileName, logFilePath);
            if (occurenceList == null || occurenceList.Count == 0)
            {
                Utility.Log("SolidEdgeOccurrenceDeleteFromExcel: " + "Occurrence List Is Empty", logFilePath);
                return;
            }
            List<String> MasterAssemblyList = MasterAssemblyReader.getComponents();

            if (MasterAssemblyList == null || MasterAssemblyList.Count == 0)
            {
                Utility.Log("SolidEdgeOccurrenceDeleteFromExcel: " + "Remove Component List is Empty", logFilePath);
                return;
            }
            var occurenceListDelete = string.Join(",", occurenceList);
            Utility.Log("occurenceList: " + occurenceListDelete.ToString(), logFilePath);
            var MasterAssemblyListDelete = string.Join(",", MasterAssemblyList);
            Utility.Log("MasterAssemblyList: " + MasterAssemblyListDelete.ToString(), logFilePath);
            var ListOfComponentsToDelete = occurenceList.Except(MasterAssemblyList).ToList();
            if (ListOfComponentsToDelete == null || ListOfComponentsToDelete.Count == 0)
            {
                Utility.Log("SolidEdgeOccurrenceDeleteFromExcel: " + "ListOfComponentsToDelete is Empty", logFilePath);
                return;
            }

            List<String> ComponentsToDelete = (List<String>)ListOfComponentsToDelete;

            if (ComponentsToDelete == null || ComponentsToDelete.Count == 0)
            {
                Utility.Log("SolidEdgeOccurrenceDeleteFromExcel: " + "ComponentsToDelete is Empty", logFilePath);
                return;
            }
            var ComponentStringDelete = string.Join(",", ComponentsToDelete);
            Utility.Log("Components to be Deleted: " + ComponentStringDelete, logFilePath);
            String SearchAssemblyFileFolder = System.IO.Path.GetDirectoryName(assemblyFileName);
            Utility.Log("SearchAssemblyFileFolder: " + SearchAssemblyFileFolder, logFilePath);

            String[] AssemblyFiles = Directory.GetFiles(SearchAssemblyFileFolder, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".asm", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();
            if (AssemblyFiles == null || AssemblyFiles.Length == 0)
            {
                Utility.Log("SearchAssemblyFileFolder -  No Assembly Files Found ", logFilePath);
                return;
            }
            List<string> asmFilesTemp = new List<string>();
            foreach (String asmFile in AssemblyFiles)
            {
                //if (Program.listOfFileNamesInSession.Contains(Path.GetFileName(asmFile.ToLower().Trim())))
                //{
                   //if (occurenceList.Contains(Path.GetFileName(asmFile.ToLower().Trim()))) {
                    Utility.Log(asmFile + " is applicable for this assembly", logFilePath);
                    asmFilesTemp.Add(asmFile);
                //}
                //else
                    //Utility.Log(asmFile + " is not applicable for this assembly", logFilePath);
            }
            AssemblyFiles = asmFilesTemp.ToArray();


            foreach (string AssemblyFile in AssemblyFiles)
            {
                InvokeSTAThread<String, String, List<String>, String>(OpenAndDeleteOccurence, AssemblyFile, assemblyFileName, ComponentsToDelete, logFilePath);

            }

            InvokeSTAThread<String, String>(OpenTopLevelAssembly, assemblyFileName, logFilePath);

        }

        [STAThread]
        private void OpenTopLevelAssembly(string assemblyFileName, string logFilePath)
        {
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;
            SolidEdgeFramework.Application objApp = null;
            try
            {
                OleMessageFilter.Register();

                objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect();
                if (objApp == null) return;
                objDocuments = objApp.Documents;
                objApp.DisplayAlerts = false;
                if (objDocuments == null) return;
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                Utility.Log("DEBUG - InputFile is Opened : " + assemblyFileName, logFilePath);

                if (objAssemblyDocument == null)
                {
                    Utility.ResetAlerts(objApp, false, logFilePath);
                    Utility.Log("DEBUG - InputFile is NOT Opened : ", logFilePath);
                    return;
                }

                if (objAssemblyDocument != null)
                {
                    Marshal.ReleaseComObject(objAssemblyDocument);
                    objAssemblyDocument = null;
                }
                Utility.ResetAlerts(objApp, false, logFilePath);

            }
            catch (Exception ex)
            {
                Utility.Log("Exception: " + ex.Message, logFilePath);
                Utility.Log("Exception: " + ex.Source, logFilePath);
                Utility.Log("Exception: " + ex.StackTrace, logFilePath);
                Utility.ResetAlerts(objApp, false, logFilePath);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }

        }

        [STAThread]
        private void OpenAndDeleteOccurence(string AssemblyFile, String assemblyFileName, List<string> ComponentsToDelete, string logFilePath)
        {
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;
            SolidEdgeFramework.Application objApp = null;
            try
            {
                OleMessageFilter.Register();

                objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect();
                //SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
                if (objApp == null)
                {
                    Utility.Log("DEBUG : Solid Edge Application Object is NULL - ", logFilePath);
                    return;
                }

                objDocuments = objApp.Documents;
                objApp.DisplayAlerts = false;
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(AssemblyFile);
                Utility.Log("DEBUG - InputFile is Opened : " + AssemblyFile, logFilePath);

                if (objAssemblyDocument == null)
                {
                    Utility.Log("DEBUG - InputFile is NOT Opened : ", logFilePath);
                    Utility.ResetAlerts(objApp, false, logFilePath);
                    return;
                }

                if (objAssemblyDocument != null)
                {
                    // This is for Top Assembly Alone
                    // Utility.Log("AssemDoc.Name : " + objAssemblyDocument.Name, logFilePath);     


                    occurrences = objAssemblyDocument.Occurrences;

                    if (occurrences == null || occurrences.Count == 0)
                    {
                        Utility.Log("occurrences is Empty--", logFilePath);
                        Utility.ResetAlerts(objApp, false, logFilePath);
                        return;
                    }

                    //for (int i = 1; i <= occurrences.Count; i++)
                    for (int i = occurrences.Count; i >= 1; i--)
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
                            Utility.Log("Cant Parse OccurenceName--", logFilePath);
                            continue;
                        }


                        //Utility.Log("-----------------------------------------", logFilePath);
                        Utility.Log("occurenceName--" + occurrence.Name, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        //Utility.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        //Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);
                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            if (ComponentsToDelete.Contains(occurenceName) == true)
                            {
                                OccurrenceDelete(occurrence, true, logFilePath);
                                SaveAssembly(objAssemblyDocument, logFilePath);
                            }
                            

                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            if (ComponentsToDelete.Contains(occurenceName) == true)
                            {
                                OccurrenceDelete(occurrence, true, logFilePath);
                                SaveAssembly(objAssemblyDocument, logFilePath);
                            }
                            

                        }
                        else if (occurrence.OccurrenceFileName.EndsWith(".asm") == true)
                        {
                            if (ComponentsToDelete.Contains(occurenceName) == true)
                            {
                                OccurrenceDelete(occurrence, true, logFilePath);
                                SaveAssembly(objAssemblyDocument, logFilePath);
                            }
                            
                        }

                    }

                    SaveAndCloseAssembly(objAssemblyDocument, logFilePath);

                    if (objApp != null) objApp.DisplayAlerts = false;
                }
                else
                {
                    Utility.ResetAlerts(objApp, false, logFilePath);
                    return;
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Exception: " + ex.Message, logFilePath);
                Utility.Log("Exception: " + ex.Source, logFilePath);
                Utility.Log("Exception: " + ex.StackTrace, logFilePath);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }
        }

        private void SaveAssembly(SolidEdgeAssembly.AssemblyDocument assemblyDoc, String logFilePath)
        {
            try
            {
                if (assemblyDoc.ReadOnly == false)
                {
                    Utility.Log("Save assembly: " + assemblyDoc.Name, logFilePath);
                    assemblyDoc.Save();
                }


            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
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
                Utility.Log(ex.Message, logFilePath);
            }

        }

        private void SaveAndCloseAssembly(SolidEdgeAssembly.AssemblyDocument assemblyDoc, String logFilePath)
        {
            try
            {
                if (assemblyDoc.ReadOnly == false)
                {
                    assemblyDoc.Save();
                }

                // 20-12-2018 Commented so that the Assembly is not Closed Entirely
                //05-01-2018 - This is an Issue, Sub Assembly Needs to be Closed. Main Assembly Should be Opened.
                if (assemblyDoc.ReadOnly == false)
                {
                    Utility.Log("close assembly: " + assemblyDoc.Name, logFilePath);
                    assemblyDoc.Close();
                }

                Marshal.ReleaseComObject(assemblyDoc);
                assemblyDoc = null;
            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
            }



        }


        private void savePart(SolidEdgePart.PartDocument partDoc, String logFilePath)
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
                Utility.Log(ex.Message, logFilePath);
            }
        }

        private void saveSheet(SolidEdgePart.SheetMetalDocument sheetMetalDoc, String logFilePath)
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
                Utility.Log(ex.Message, logFilePath);
            }
        }

        private void ReadAllOccurrencesFromAssembly(String assemblyFileName, String logFilePath)
        {
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;
            SolidEdgeFramework.Application objApp = null;

            try
            {
                OleMessageFilter.Register();

                objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect();
                //SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
                if (objApp == null)
                {
                    Utility.Log("DEBUG : Solid Edge Application Object is NULL - ", logFilePath);
                    return;
                }

                objDocuments = objApp.Documents;
                objApp.DisplayAlerts = false;

                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;
                Utility.Log("DEBUG - InputFile is Opened : " + objAssemblyDocument.FullName, logFilePath);

                if (objAssemblyDocument == null)
                {
                    Utility.ResetAlerts(objApp, false, logFilePath);
                    Utility.Log("DEBUG - InputFile is NOT Opened : ", logFilePath);
                    return;
                }

                if (objAssemblyDocument != null)
                {
                    // This is for Top Assembly Alone
                    // Utility.Log("AssemDoc.Name : " + objAssemblyDocument.Name, logFilePath);                    
                    occurrences = objAssemblyDocument.Occurrences;

                    if (occurrences == null || occurrences.Count == 0)
                    {
                        Utility.ResetAlerts(objApp, false, logFilePath);
                        Utility.Log("occurrences is Empty--", logFilePath);
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
                            Utility.Log("Cant Parse OccurenceName--", logFilePath);
                            continue;
                        }


                        if (occurenceList.Contains(occurenceName) == true)
                        {
                            continue;
                        }
                        else
                        {
                            occurenceList.Add(occurenceName);
                        }

                        //Utility.Log("-----------------------------------------", logFilePath);
                        //Utility.Log("occurenceName--" + occurenceName, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        //Utility.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        //Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);
                        if (occurrence.OccurrenceFileName.EndsWith(".asm") == true)
                        {
                            SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;
                            assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;

                            traverseAssembly(assemDoc1, logFilePath);

                        }

                        //Utility.Log("-----------------------------------------", logFilePath);


                    }
                    //SaveAndCloseAssembly(objAssemblyDocument, logFilePath);
                    SaveAssembly(objAssemblyDocument, logFilePath);
                    Utility.ResetAlerts(objApp, false, logFilePath);
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Exception: " + ex.Message, logFilePath);
                Utility.Log("Exception: " + ex.Source, logFilePath);
                Utility.Log("Exception: " + ex.StackTrace, logFilePath);
                Utility.ResetAlerts(objApp, false, logFilePath);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }

        }

        private void traverseAssembly(SolidEdgeAssembly.AssemblyDocument assemDoc, String logFilePath)
        {
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            if (assemDoc == null)
            {
                Utility.Log("assemDoc is Empty: " + assemDoc.Name, logFilePath);
                return;
            }

            occurrences = assemDoc.Occurrences;
            //Utility.Log("occurrences.Count: " + occurrences.Count, logFilePath);
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
                    Utility.Log("Cant Parse OccurenceName--", logFilePath);
                    continue;
                }
                if (occurenceList.Contains(occurenceName) == true)
                {
                    continue;
                }
                else
                {
                    occurenceList.Add(occurenceName);
                }


                //Utility.Log("-----------------------------------------", logFilePath);
                //Utility.Log("occurenceName--" + occurenceName, logFilePath);
                int occurenceQty = occurrence.Quantity;
                //Utility.Log("occurenceQty--" + occurenceQty, logFilePath);
                String ocurenceFileName = occurrence.OccurrenceFileName;
                //Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {   
                        SolidEdgeAssembly.AssemblyDocument assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                        traverseAssembly(assemDoc1, logFilePath);
                    
                }
                
            }

        }
    }
}
