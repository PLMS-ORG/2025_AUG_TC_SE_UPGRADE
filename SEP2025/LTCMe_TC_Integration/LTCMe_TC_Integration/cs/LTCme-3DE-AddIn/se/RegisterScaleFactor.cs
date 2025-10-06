using DemoAddInTC.utils;
using SolidEdgeCommunity;
using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DemoAddInTC.se
{
    class RegisterScaleFactor
    {
        public static Dictionary<string, Dictionary<string, double>> scaleFactorDictionary = new Dictionary<string, Dictionary<string, double>>();
        static RegisterScaleFactor()
        {
            
        }

        public RegisterScaleFactor()
        {
        }

        private static void addInformationToDictionary(string draftFileFULLPath, string viewName, double maxScaleLength, string logFilePath)
        {
            string draftFileName = Path.GetFileName(draftFileFULLPath);
            if (RegisterScaleFactor.scaleFactorDictionary.ContainsKey(draftFileName))
            {
                Dictionary<string, double> viewScaleFactorDictionary = null;
                RegisterScaleFactor.scaleFactorDictionary.TryGetValue(draftFileName, out viewScaleFactorDictionary);
                if (!viewScaleFactorDictionary.ContainsKey(viewName))
                {
                    viewScaleFactorDictionary.Add(viewName, maxScaleLength);
                }
            }
            else
            {
                Dictionary<string, double> viewScaleFactorDictionary = new Dictionary<string, double>()
                {
                    { viewName, maxScaleLength }
                };
                RegisterScaleFactor.scaleFactorDictionary.Add(draftFileName, viewScaleFactorDictionary);
            }
        }

        public static bool registerScaleFactor(string DraftFileFULLPath, string logFilePath)
        {
            bool flag;
            Utlity.Log(string.Concat("RegisterScaleFactor -  DraftFileFULLPath: ", DraftFileFULLPath), logFilePath, null);
            Documents objDocuments = null;
            Application objApp = SolidEdgeUtils.Connect();
            if (objApp != null)
            {
                DraftDocument objDraftDocument = null;
                objDocuments = objApp.Documents;
                if (objDocuments != null)
                {
                    try
                    {
                        OleMessageFilter.Register();
                    }
                    catch (Exception exception)
                    {
                        Exception ex = exception;
                        Utlity.Log(string.Concat("DEBUG - OleMessageFilter Register : ", DraftFileFULLPath), logFilePath, null);
                        Utlity.Log(ex.Message, logFilePath, null);
                    }
                    try
                    {
                        if (!File.Exists(DraftFileFULLPath))
                        {
                            Utlity.ResetAlerts(objApp, true, logFilePath);
                            Utlity.Log(string.Concat("Draft File Does not Exist ", DraftFileFULLPath), logFilePath, null);
                            flag = false;
                            return flag;
                        }
                        else
                        {
                            objApp.SetGlobalParameter(SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalOpenAsReadOnlyDftFile, false);
                            objApp.DisplayAlerts = false;
                            objDraftDocument = (DraftDocument)objDocuments.Open(DraftFileFULLPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            if (objDraftDocument.ReadOnly == true)
                            {
                                bool WriteAccess = false;
                                objDraftDocument.SeekWriteAccess(out WriteAccess);
                                if (!WriteAccess)
                                {
                                    Utlity.ResetAlerts(objApp, true, logFilePath);
                                    Utlity.Log(string.Concat("Could not get WriteAccess to--", DraftFileFULLPath), logFilePath, null);
                                    flag = false;
                                    return flag;
                                }
                            }
                        }
                    }
                    catch (Exception exception1)
                    {
                        Exception ex = exception1;
                        Utlity.ResetAlerts(objApp, true, logFilePath);
                        Utlity.Log(string.Concat("DEBUG - Unable to Open : ", DraftFileFULLPath), logFilePath, null);
                        Utlity.Log(ex.Message, logFilePath, null);
                        flag = false;
                        return flag;
                    }
                    if (objDraftDocument == null)
                    {
                        flag = false;
                    }
                    else
                    {
                        Utlity.Log("------ PROCESS SHEETS ------", logFilePath, null);
                        SolidEdgeDraft.Sheets sheets = objDraftDocument.Sheets;
                        SolidEdgeDraft.Sections sections = objDraftDocument.Sections;
                        SolidEdgeDraft.Sheet ActiveSheet = objDraftDocument.ActiveSheet;
                        foreach (SheetGroup sheetGroup in (SheetGroups)objDraftDocument.SheetGroups)
                        {
                            foreach (Sheet sheet in sheetGroup.Sheets)
                            {
                                foreach (DrawingView drawingView in sheet.DrawingViews)
                                {
                                    double scalefactor = drawingView.ScaleFactor;
                                    double maxScaleLength = 0;
                                    SolidEdgeDraft.DraftProfile croppingBoundaryProfile = drawingView.CroppingBoundaryProfile;
                                    foreach (Line2d line in drawingView.CroppingBoundaryProfile.Lines2d)
                                    {
                                        if (scalefactor * line.Length > maxScaleLength) maxScaleLength = scalefactor * line.Length;
                                    }
                                    try
                                    {
                                        Utility.Log(string.Concat(drawingView.Name, "=", maxScaleLength.ToString()), logFilePath);
                                    }
                                    catch (Exception ex)
                                    {
                                        Utility.Log("Drawing View Name Exception: " + ex.Message, logFilePath);
                                        continue;
                                    }
                                    Utility.Log(string.Concat("drawingView Name: =", drawingView.Name), logFilePath);
                                    Utility.Log(string.Concat("maxScaleLength: =", maxScaleLength.ToString()), logFilePath);
                                    RegisterScaleFactor.addInformationToDictionary(DraftFileFULLPath, drawingView.Name, maxScaleLength, logFilePath);
                                    Utility.Log("Setting description: ", logFilePath);
                                    drawingView.Description = maxScaleLength.ToString();
                                }
                            }
                        }
                        try
                        {
                            try
                            {
                                if (objDraftDocument.ReadOnly == false)
                                {
                                    Utility.Log("DEBUG - Close & Save", logFilePath);
                                    objDraftDocument.Close(true, Type.Missing, Type.Missing);
                                }
                                Marshal.ReleaseComObject(objDraftDocument);
                                objDraftDocument = null;
                                if (objApp != null)
                                {
                                    objApp.DisplayAlerts = true;
                                }
                            }
                            catch (Exception exception2)
                            {
                                Exception ex = exception2;
                                Utlity.ResetAlerts(objApp, true, logFilePath);
                                Utlity.Log(string.Concat("DEBUG - Unable to Save the Document : ", ex.Message), logFilePath, null);
                                flag = false;
                                return flag;
                            }
                        }
                        finally
                        {
                            OleMessageFilter.Unregister();
                        }
                        flag = true;
                    }
                }
                else
                {
                    Utlity.ResetAlerts(objApp, true, logFilePath);
                    Utlity.Log(string.Concat("DEBUG -  objDocuments is NULL : ", DraftFileFULLPath), logFilePath, null);
                    flag = false;
                }
            }
            else
            {
                Utlity.Log(string.Concat("DEBUG -  objApp is NULL : ", DraftFileFULLPath), logFilePath, null);
                flag = false;
            }
            return flag;
        }


    }
}
