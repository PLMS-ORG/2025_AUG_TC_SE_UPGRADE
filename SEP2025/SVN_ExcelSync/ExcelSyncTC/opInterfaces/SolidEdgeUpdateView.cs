using ExcelSyncTC.utils;
using SolidEdgeCommunity;
using SolidEdgeDraft;
using SolidEdgeFrameworkSupport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace ExcelSyncTC.opInterfaces
{
    class SolidEdgeUpdateView
    {
        public static void SearchDraftFile(String assemblyFileName, String logFilePath)
        {
            String searchDrawingsFolder = System.IO.Path.GetDirectoryName(assemblyFileName);
            Utlity.Log("searchDrawingsFolder: " + searchDrawingsFolder, logFilePath);

            String[] draftFiles = Directory.GetFiles(searchDrawingsFolder, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".dft", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();
            if (draftFiles == null || draftFiles.Length == 0)
            {
                Utlity.Log("SearchDraftFile -  No Draft Files Found ", logFilePath);
                return;
            }

            //Removing drafts not applicable for this assy
            List<string> draftFilesTemp = new List<string>();
            foreach (String draftFile in draftFiles)
            {
                if (SyncDwg.listOfFileNamesInSession.Contains(Path.GetFileName(draftFile.ToLower().Trim())))
                {
                    Utlity.Log(draftFile + " is applicable for this assembly", logFilePath);
                    draftFilesTemp.Add(draftFile);
                }
                else
                    Utlity.Log(draftFile + " is not applicable for this assembly", logFilePath);
            }
                draftFiles = draftFilesTemp.ToArray();

            foreach (String draftFile in draftFiles)
            {
                try
                {
                    Utlity.Log("findOutOfDateDrawing: " + System.DateTime.Now.ToString(), logFilePath);
                    Thread myThread = new Thread(() => findOutOfDateDrawing(draftFile, logFilePath));
                    myThread.SetApartmentState(ApartmentState.STA);
                    myThread.Start();
                    myThread.Join();
                }
                catch (Exception ex)
                {
                    Utlity.Log("SolidEdgeUpdateView, findOutOfDateDrawing: " + ex.Message, logFilePath);
                    return;
                }
            }


        }
        [STAThread]
        public static bool findOutOfDateDrawing(String DraftFileFULLPath, String logFilePath)
        {
            Utlity.Log("findOutOfDateDrawing -  DraftFileFULLPath: ", DraftFileFULLPath);
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect();
            if (objApp == null)
            {
                Utlity.Log("DEBUG -  objApp is NULL : " + DraftFileFULLPath, logFilePath);
                return false;
            }

            SolidEdgeDraft.DraftDocument objDraftDocument = null;
            objDocuments = objApp.Documents;

            if (objDocuments == null)
            {
                Utlity.ResetAlerts(objApp, true, logFilePath);
                Utlity.Log("DEBUG -  objDocuments is NULL : " + DraftFileFULLPath, logFilePath);
                return false;
            }

            try
            {
                OleMessageFilter.Register();

            }
            catch (Exception ex)
            {
                Utlity.Log("DEBUG - OleMessageFilter Register : " + DraftFileFULLPath, logFilePath);
                Utlity.Log(ex.Message, logFilePath);
            }


            try
            {
                if (System.IO.File.Exists(DraftFileFULLPath) == true)
                {
                    objApp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalOpenAsReadOnlyDftFile, false);
                    objApp.DisplayAlerts = false;
                    objDraftDocument = (SolidEdgeDraft.DraftDocument)objDocuments.Open(DraftFileFULLPath);
                    if (objDraftDocument.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        objDraftDocument.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utlity.ResetAlerts(objApp, true, logFilePath);
                            Utlity.Log("Could not get WriteAccess to--" + DraftFileFULLPath, logFilePath);
                            return false;
                        }
                    }
                }
                else
                {
                    Utlity.ResetAlerts(objApp, true, logFilePath);
                    Utlity.Log("Draft File Does not Exist " + DraftFileFULLPath, logFilePath);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Utlity.ResetAlerts(objApp, true, logFilePath);
                Utlity.Log("DEBUG - Unable to Open : " + DraftFileFULLPath, logFilePath);
                Utlity.Log(ex.Message, logFilePath);
                return false;
            }

            //Make sure we have a document.
            if (objDraftDocument != null)
            {
                // 26-10-2023: Parts List Update
                SolidEdgeDraft.PartsLists partsList = objDraftDocument.PartsLists;


                if (partsList != null && partsList.Count > 0)
                {
                    // 27-11-2024 | Murali | Update the Parts List of each Draft file
                    updatePartsList(partsList, logFilePath);
                }
                else
                {
                    Utlity.Log("DEBUG - No partsList in the Draft File : " + DraftFileFULLPath, logFilePath);
                }

                Utlity.Log("------ PROCESS SHEETS ------", logFilePath);
                // Get a reference to the sheets collection.
                SolidEdgeDraft.Sheets sheets = objDraftDocument.Sheets;
                SolidEdgeDraft.Sections sections = objDraftDocument.Sections;
                SolidEdgeDraft.Sheet ActiveSheet = objDraftDocument.ActiveSheet;

                foreach (SolidEdgeDraft.Sheet sheet in sections.WorkingSection.Sheets)
                {
                    SolidEdgeDraft.DrawingViews views = sheet.DrawingViews;

                    if (views == null)
                    {
                        Marshal.ReleaseComObject(sheet);
                        Utlity.Log("No Views in the Sheet: " + sheet.Name, logFilePath);
                        continue;
                    }
                    Utlity.Log("Sheet: " + sheet.Name, logFilePath);
                    foreach (SolidEdgeDraft.DrawingView vw in views)
                    {
                        Utlity.Log("View Name " + vw.Name + "::::" + " View UpToDateStatus " + vw.IsUpToDate, logFilePath);
                        if (vw.IsUpToDate == false)
                        {
                            //Utlity.Log("DraftFileName: " + DraftFileFULLPath + "  Sheet: " + sheet.Name +
                            //" View Name " + vw.Name + " View UpToDateStatus " + vw.IsUpToDate, outPutFilePath);
                            //Utlity.Log("DraftFileName: " + SolidEdgeDraftFilePath + "  Sheet: " + sheet.Name                            , outPutFilePath);

                            Utlity.Log("ForceUpdate: " + sheet.Name, logFilePath);
                            try
                            {
                                vw.ForceUpdate();
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("ForceUpdate: " + "Exception: " + ex.Message, logFilePath);
                            }
                        }
                        if (vw != null) Marshal.ReleaseComObject(vw);
                    }

                    if (views != null)
                    {
                        Marshal.ReleaseComObject(views);
                        views = null;
                    }
                    Marshal.ReleaseComObject(sheet);
                }





                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                    sheets = null;
                }

                if (sections != null)
                {
                    Marshal.ReleaseComObject(sections);
                    sections = null;
                }


            }
            else
            {
                Utlity.ResetAlerts(objApp, true, logFilePath);
                Utlity.Log("Replacing Exception: ", logFilePath);
                throw new System.Exception("No active document.");
            }

            // ==== Autoscaling Implementation in Drawing Views == Murali - Implemented on February 16 2022 - Implemented on February 16 2022 but Added to this workspace on 6-DEC-2024
            try
            {
                foreach (SheetGroup sheetGroup in (SheetGroups)objDraftDocument.SheetGroups)
                {
                    Utlity.Log("DEBUG - Found Sheet Group: ", logFilePath);
                    foreach (Sheet sheet1 in sheetGroup.Sheets)
                    {
                        Utlity.Log(string.Concat("DEBUG - Found Sheet : ", sheet1.Name), logFilePath);
                        foreach (DrawingView drawingView1 in sheet1.DrawingViews)
                        {

                            double length = 0;
                            foreach (Line2d lines2d in drawingView1.CroppingBoundaryProfile.Lines2d)
                            {
                                if (lines2d.Length > length)
                                {
                                    length = lines2d.Length;
                                }
                            }
                            string scaleLength = "";
                            string description = drawingView1.Description;
                            Utlity.Log(string.Concat("DEBUG - Found Sheet : ", sheet1.Name), logFilePath);
                            Utlity.Log(string.Concat("DEBUG - description : ", description), logFilePath);
                            if ((description == null || description.Equals("") == true))
                            {
                                try
                                {
                                    Utlity.Log(string.Concat("View ForceUpdate : ", drawingView1.Name), logFilePath);
                                    drawingView1.ForceUpdate();
                                }
                                catch (Exception ex)
                                {
                                    Utlity.Log(string.Concat("ForceUpdate: Exception: ", ex.Message), logFilePath);
                                }
                            }
                            else
                            {
                                string[] strArrays = description.Split(new char[] { ':' });
                                if ((strArrays != null && strArrays.Length == 2))
                                {
                                    scaleLength = strArrays[0]; // scalefactor value
                                    string str1 = strArrays[1]; // whether to autoscale or not
                                    double num = 0;
                                    try
                                    {
                                        Utlity.Log(string.Concat("DEBUG - strMaxScaleLength : ", scaleLength), logFilePath);
                                        double.TryParse(scaleLength, out num);
                                    }
                                    catch (Exception ex11)
                                    {

                                        Utlity.Log(string.Concat("DEBUG - Exception : ", ex11.Message), logFilePath);
                                        continue;
                                    }
                                    if (num != 0)
                                    {
                                        Utlity.Log(string.Concat("drawingView.Name: ", drawingView1.Name), logFilePath);
                                        if ((str1 != null && str1.Equals("") == false))
                                        {
                                            if (str1.Equals("AUTOSCALE", StringComparison.OrdinalIgnoreCase))
                                            {
                                                Utlity.Log(string.Concat("maxScaleLength : ", num.ToString()), logFilePath);
                                                Utlity.Log(string.Concat("maxLength : ", length.ToString()), logFilePath);
                                                if (length != 0)
                                                {
                                                    drawingView1.ScaleFactor = (num / length);
                                                    double scaleFactor = drawingView1.ScaleFactor;
                                                    Utlity.Log(string.Concat("ScaleFactor : ", scaleFactor.ToString()), logFilePath);
                                                }
                                                else
                                                {
                                                    Utlity.Log("DEBUG - Unable to perform autoScale : " + drawingView1.Name, logFilePath);
                                                    continue;

                                                }
                                            }
                                            else
                                            {
                                                Utlity.Log("DEBUG - autoScaleOption is not enabled : " + drawingView1.Name, logFilePath);
                                                continue;
                                            }

                                            try
                                            {
                                                Utlity.Log(string.Concat("View ForceUpdate : ", drawingView1.Name), logFilePath);
                                                drawingView1.ForceUpdate();
                                            }
                                            catch (Exception ex1)
                                            {

                                                Utlity.Log(string.Concat("ForceUpdate: Exception: ", ex1.Message), logFilePath);
                                            }
                                        }
                                        else
                                        {
                                            Utlity.Log("DEBUG - autoScaleOption is Empty : ", logFilePath);
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        Utlity.Log("DEBUG - Scale length is zero : " + drawingView1.Name, logFilePath);
                                        continue;
                                    }
                                }
                                else
                                {
                                    Utlity.Log("DEBUG - autoScaleOption is not enabled for view : " + drawingView1.Name, logFilePath);
                                    continue;

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex3)
            {
                Utlity.Log(string.Concat("DEBUG - Exception : ", ex3.Message), logFilePath);

            }
            // ==== Autoscaling Implementation in Drawing Views == Murali - Implemented on February 16 2022 but Added to this workspace on 6-DEC-2024

            try
            {
                if (objDraftDocument.ReadOnly == false)
                {
                    Utlity.Log("DEBUG - ClearAllDimensionTrackerEntries", logFilePath);
                    objDraftDocument.ClearAllDimensionTrackerEntries();

                    Utlity.Log("DEBUG - Close & Save", logFilePath);
                    objDraftDocument.Close(true);
                }

                Marshal.ReleaseComObject(objDraftDocument);
                //Marshal.FinalReleaseComObject(objApp.Documents);
                objDraftDocument = null;
                if (objApp != null) objApp.DisplayAlerts = true;  

            }
            catch (Exception ex)
            {
                Utlity.ResetAlerts(objApp, true, logFilePath);
                Utlity.Log("DEBUG - Unable to Save the Document : " + ex.Message, logFilePath);
                return false;
            }

            finally
            {
                OleMessageFilter.Unregister();
            }
           

            return true;
        }

        // 27-11-2024 | Murali | Update the Parts List of each Draft file
        private static void updatePartsList(PartsLists partsLists, string logFilePath)
        {
            foreach (SolidEdgeDraft.PartsList partsList in partsLists)
            {
                if (partsList.IsUpToDate == false)
                {
                    Utlity.Log("DEBUG - partsList Update: ", logFilePath);
                    partsList.Update();
                }
                else
                {
                    Utlity.Log("DEBUG - partsList is already Up to Date.. ", logFilePath);
                }
            }
        }
    }
}
