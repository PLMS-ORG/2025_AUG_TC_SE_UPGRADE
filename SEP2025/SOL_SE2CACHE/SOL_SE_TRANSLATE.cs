/*
 * 15/03/2019 -Murali- Modified code to get document object instead of part object 
 **/
using SolidEdgeCommunity;
//using SolidEdgeCommunity.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{    

    class SOL_SE_TRANSLATE : IsolatedTaskProxy
    {
        string itemIDTempStorage = null;
        public string m_parentType = "";

        public void SaveDraftAs1STATThread(String InputFile,
            String logFilePath, String OutputFolder, String itemID, String parentType)
        {
            itemIDTempStorage = itemID;
            m_parentType = parentType;
            InvokeSTAThread<String, String, String, String>(SaveDraftAs1, InputFile, logFilePath, OutputFolder, itemID);
        }

        [STAThread]
        private void SaveDraftAs1( String InputFile, String logFilePath, String OutputFolder,String itemID)
        {
            String Format1 = "PDF";
            //String Format2 = "DXF";

            String StageDir = Path.GetDirectoryName(InputFile);

            SolidEdge.Framework.Interop.Documents objDocuments = null;
            //SolidEdge.Framework.Interop.Application objApp = SE_SESSION.getSolidEdgeSession();
            //objApp.DoIdle();
            SolidEdgeDraft.DraftDocument objDraftDocument = null;


            String draftFileFULLPath = InputFile; // Path.Combine(StageDir, InputFile);
            try
            {
                Utility.Log("DEBUG - Opening : " + draftFileFULLPath, logFilePath);
                OleMessageFilter.Register();
                SolidEdge.Framework.Interop.Application objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Start();
                if (objApp == null) return;
                objDocuments = objApp.Documents;
                if (objDocuments == null) return;
                if (System.IO.File.Exists(draftFileFULLPath) == true)
                {
                    objApp.DisplayAlerts = false;
                    objDraftDocument = objDocuments.Open(draftFileFULLPath);

                    if (objDraftDocument.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        objDraftDocument.SeekWriteAccess(out WriteAccess);
                       
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + draftFileFULLPath, logFilePath);
                            return ;
                        }
                    }

                    Utility.Log("DEBUG - InputFile is Opened : " + draftFileFULLPath, logFilePath);
                }
                else
                {
                    Utility.Log("Draft File Does not Exist " + draftFileFULLPath, logFilePath);
                    return ;
                }

                String outputFile1 = Path.ChangeExtension(InputFile, "." + Format1);
                outputFile1 = Path.GetFileName(outputFile1);
                Utility.Log("DEBUG - outputFile : " + outputFile1, logFilePath);
                String outPutFileFULLPath1 = Path.Combine(OutputFolder, outputFile1);

                //String outputFile2 = Path.ChangeExtension(InputFile, "." + Format2);
                //outputFile2 = Path.GetFileName(outputFile2);
                //Utility.Log("DEBUG - outputFile : " + outputFile2, logFilePath);
                //String outPutFileFULLPath2 = Path.Combine(OutputFolder, outputFile2);

                if (objDraftDocument != null)
                {
                    ActivateWorkingSheet(objDraftDocument,logFilePath);

                    try
                    {

                        //objDraftDocument.SaveAs(outputFile, false, OutPutFileFormat, false, false);
                        if (System.IO.File.Exists(outPutFileFULLPath1) == true)
                        {
                            Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath1, logFilePath);
                            System.IO.File.Delete(outPutFileFULLPath1);
                        }
                        Utility.Log("DEBUG - SaveAs: " + outPutFileFULLPath1, logFilePath);
                        if (Format1.Equals("PDF", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            setPDFSaveConstants(objApp);
                        }
                       
                        objDraftDocument.SaveAs(outPutFileFULLPath1);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath1, logFilePath);
                        Utility.Log("Unable to Save " + Format1 + " for " + itemIDTempStorage, Program.failureFilePath);
                        Utility.Log(ex.Message, logFilePath);
                        return ;
                    }
                    if (File.Exists(outPutFileFULLPath1) == false)
                    {
                        Utility.Log("DEBUG - SaveAs to PDF Failed : " + outPutFileFULLPath1, logFilePath);
                        return ;
                    }

                    // Format2 --- DXF
                    //if (m_parentType.Equals("SHEETMETAL", StringComparison.OrdinalIgnoreCase) == true)
                    //{
                    //    Utility.Log("DEBUG - DXF WILL BE SAVED ONLY FOR SHEETMETAL DRAFTS..", logFilePath);
                    //    try
                    //    {

                    //        //objDraftDocument.SaveAs(outputFile, false, OutPutFileFormat, false, false);
                    //        if (System.IO.File.Exists(outPutFileFULLPath2) == true)
                    //        {
                    //            Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath2, logFilePath);
                    //            System.IO.File.Delete(outPutFileFULLPath2);
                    //        }
                    //        Utility.Log("DEBUG - SaveAs: " + outPutFileFULLPath2, logFilePath);
                    //        if (Format2.Equals("DXF", StringComparison.OrdinalIgnoreCase) == true)
                    //        {
                    //            //setPDFSaveConstants(objApp);
                    //        }

                    //        objDraftDocument.SaveAs(outPutFileFULLPath2);
                            
                            
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath2, logFilePath);
                    //        Utility.Log("Unable to Save " + Format2 + " for " + itemIDTempStorage, Program.failureFilePath);
                    //        Utility.Log(ex.Message, logFilePath);
                    //        return;
                    //    }
                    //    if (File.Exists(outPutFileFULLPath2) == false)
                    //    {
                    //        Utility.Log("DEBUG - SaveAs to DXF Failed : " + outPutFileFULLPath2, logFilePath);
                    //        return;
                    //    }
                    //}

                    try
                    {
                        if (objDraftDocument.ReadOnly == false)
                        {
                            Utility.Log("DEBUG - Save", logFilePath);
                            objDraftDocument.Save();
                        }

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Save the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to Save the draft file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }
                    try
                    {
                        Utility.Log("DEBUG - Close", logFilePath);
                        objDraftDocument.Close(true);

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Close the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to Close the draft of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }

                    Marshal.ReleaseComObject(objDraftDocument);
                    //Marshal.FinalReleaseComObject(objApp.Documents);
                    objDraftDocument = null;
                    Utility.Log("DEBUG - Quit : " + draftFileFULLPath, logFilePath);
                    objApp.Quit();
                    OleMessageFilter.Unregister();
                    
                }

            }
            catch (Exception ex)
            {
                Utility.Log("DEBUG - Unable to Open : " + draftFileFULLPath, logFilePath);
                Utility.Log("Unable to Open draft of " + itemIDTempStorage, Program.failureFilePath);
                Utility.Log(ex.Message, logFilePath);
                return ;
            }

            return ;
        }

        // 12-2-2019, If Active Sheet is not Working Sheet and its a Background Sheet, Change it
        private void ActivateWorkingSheet(SolidEdgeDraft.DraftDocument objDraftDocument, String logFilePath)
        {
            //Make sure we have a document.
            if (objDraftDocument != null)
            {
               
                // Get a reference to the sheets collection.
                SolidEdgeDraft.Sheets sheets = objDraftDocument.Sheets;
                SolidEdgeDraft.Sections sections = objDraftDocument.Sections;
                SolidEdgeDraft.Sheet ActiveSheet = objDraftDocument.ActiveSheet;

                if (ActiveSheet == null) return;

                if (ActiveSheet.SectionType == SolidEdgeDraft.SheetSectionTypeConstants.igBackgroundSection)
                {
                    Utility.Log("------ ACTIVATING WORKING SHEET------", logFilePath);
                    foreach (SolidEdgeDraft.Sheet sheet in sections.WorkingSection.Sheets)
                    {
                        Utility.Log("DEBUG - SectionType: " + sheet.SectionType, logFilePath);
                        Utility.Log("DEBUG - Name: " + sheet.Name, logFilePath);
                        sheet.Activate();
                        Marshal.ReleaseComObject(sheet);
                        break;

                    }
                }
                //else if (ActiveSheet.SectionType == SolidEdgeDraft.SheetSectionTypeConstants.igWorkingSection)
                //{
                    
                    // 193700 - In case background sheet is visible, but not active, then there is a need to iterate and make them hidden.
                    // if this is not fixed, then empty pdf will be seen
                    //foreach (SolidEdgeDraft.Sheet sheet1 in sections.BackgroundSection.Sheets)
                    //{

                    //    if (sheet1.Visible == true)
                    //    {
                    //        Utility.Log("DEBUG - SectionType: " + sheet1.SectionType, logFilePath);
                    //        Utility.Log("DEBUG - Name: " + sheet1.Name, logFilePath);
                    //        Utility.Log("DEBUG - Visible: " + "true", logFilePath);
                    //        sections.BackgroundSection.Deactivate();
                    //    }
                    //    Marshal.ReleaseComObject(sheet1);

                    //}

                    
                    
                //}

                // 193700 - In case background sheet is visible, but not active, then there is a need to iterate and make them hidden.
                // if this is not fixed, then empty pdf will be seen
                try
                {

                    sections.BackgroundSection.Deactivate();
                }
                catch (Exception ex)
                {
                    Utility.Log("DEBUG - BackgroundSection is not deactivated, May be its already deactive ", logFilePath);
                }

                Marshal.ReleaseComObject(ActiveSheet);
                ActiveSheet = null;

                Marshal.ReleaseComObject(sections);
                sections = null;

                Marshal.ReleaseComObject(sheets);
                sheets = null;
            }
        }

        private static void setPDFSaveConstants(SolidEdge.Framework.Interop.Application objApplication)
        {

            objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSaveAllColorsBlack, false);

            objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFIncludeGridDisplay, false);

            objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFTransparentDVBackgrounds, true);

            objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFUseIndividualSheetSizes, true);

            objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetOptions, SolidEdgeConstants.DraftSaveAsPDFSheetOptionsConstants.seDraftSaveAsPDFSheetOptionsConstantsAllSheets);


            objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFPrintQualityDPI, SolidEdgeConstants.DraftSaveAsPDFPrintQualityDPIConstants.seDraftSaveAsPDFPrintQualityDPIConstants_1200);


            //objApplication.SetGlobalParameter(SolidEdge.Framework.Interop.ApplicationGlobalConstants.seApplicationGlobalDraftSaveAsPDFSheetsRange, "2222");
            

            //objApplication.SetGlobalParameter(SolidEdgeConstants.SheetMetalCommandConstants.SheetMetalSheetMetalFeaturesSaveAsFlat);

        }

        public void SavePartAsSTATThread(String StageDir, String InputFile, String Format, String logFilePath, String OutputFolder, string itemID)
        {
            itemIDTempStorage = itemID;
            InvokeSTAThread<String, String, String, String>(SavePartAs, InputFile, Format, logFilePath, OutputFolder);
        }

        
        [STAThread]
        public void SavePartAs(String InputFile, String Format, String logFilePath,String OutputFolder)
        {
            String StageDir= Path.GetDirectoryName(InputFile);
            SolidEdge.Framework.Interop.Documents objDocuments = null;
           
            
            //SolidEdgePart.PartDocument  objPartDocument = null;
            SolidEdgeFramework.SolidEdgeDocument document = null;
            String partFileFULLPath = InputFile; //Path.Combine(StageDir, InputFile);
            try
            {
                OleMessageFilter.Register();
                SolidEdge.Framework.Interop.Application objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Start();
                if (objApp == null) return;
                objDocuments = objApp.Documents;
                if (objDocuments == null) return;
                if (System.IO.File.Exists(partFileFULLPath) == true)
                {
                    objApp.DisplayAlerts = false;
                    //objPartDocument = objDocuments.Open(partFileFULLPath);
                    document = (SolidEdgeFramework.SolidEdgeDocument)objDocuments.Open(partFileFULLPath);
                    if (document.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        document.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + partFileFULLPath, logFilePath);
                            return ;
                        }
                    }

                    Utility.Log("DEBUG - InputFile is Opened : " + partFileFULLPath, logFilePath);
                }
                else
                {
                    Utility.Log("Part File Does not Exist " + partFileFULLPath, logFilePath);
                    return ;
                }

                String outputFile = Path.ChangeExtension(InputFile, "." + Format);
                outputFile = Path.GetFileName(outputFile);
                Utility.Log("DEBUG - outputFile : " + outputFile, logFilePath);
                String outPutFileFULLPath = Path.Combine(OutputFolder, outputFile);

                if (document != null)
                {
                    try
                    {

                        //objDraftDocument.SaveAs(outputFile, false, OutPutFileFormat, false, false);
                        if (System.IO.File.Exists(outPutFileFULLPath) == true)
                        {
                            Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath, logFilePath);
                            System.IO.File.Delete(outPutFileFULLPath);
                        }
                        Utility.Log("DEBUG - SaveAs: " + outPutFileFULLPath, logFilePath);
                        document.SaveAs(outPutFileFULLPath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath, logFilePath);
                        Utility.Log("Unable to Save " + Format + " for " + itemIDTempStorage, Program.failureFilePath);
                        Utility.Log(ex.Message, logFilePath);
                        return ;
                    }
                    if (File.Exists(outPutFileFULLPath) == false)
                    {
                        Utility.Log("DEBUG - SaveAs Failed : " + outPutFileFULLPath, logFilePath);
                        return ;
                    }

                    try
                    {
                        if (document.ReadOnly == false)
                        {
                            Utility.Log("DEBUG - Save", logFilePath);
                            document.Save();
                        }

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Save the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to Save part file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }
                    try
                    {
                        Utility.Log("DEBUG - Close", logFilePath);
                        document.Close(true);

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Close the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to close part file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }


                    Marshal.ReleaseComObject(document);
                    //Marshal.FinalReleaseComObject(objApp.Documents);
                    document = null;
                    Utility.Log("DEBUG - Quit ", logFilePath);
                    objApp.Quit();
                    OleMessageFilter.Unregister();
                   
                }
                Utility.Log("SOL_SE_TRANSLATE: part Save As Completed @ " + System.DateTime.Now.ToString(), logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("DEBUG - Unable to Open : " + partFileFULLPath, logFilePath);
                Utility.Log("Unable to open part file of " + itemIDTempStorage, Program.failureFilePath);
                Utility.Log(ex.Message, logFilePath);
                return ;
            }

            return ;
        }

        public void SaveAssemblyAsSTATThread(String InputFile, String Format, String logFilePath, String OutputFolder, string itemID)
        {
            itemIDTempStorage = itemID;
            InvokeSTAThread<String, String, String, String>(SaveAssemblyAs, InputFile, Format, logFilePath, OutputFolder);
        }

        [STAThread]
        public void SaveAssemblyAs(String InputFile, String Format, String logFilePath, String OutputFolder)
        {
            SolidEdge.Framework.Interop.Documents objDocuments = null;
            SolidEdge.Assembly.Interop.AssemblyDocument objAssemblyDocument = null;
            
            String StageDir = Path.GetDirectoryName(InputFile);
            String asmFileFULLPath = InputFile; // Path.Combine(StageDir, InputFile);
            try
            {
                OleMessageFilter.Register();
                SolidEdge.Framework.Interop.Application objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Start();
                if (objApp == null) return;
                objDocuments = objApp.Documents;
                if (objDocuments == null) return;
                objDocuments = objApp.Documents;
                if (System.IO.File.Exists(asmFileFULLPath) == true)
                {
                    objApp.DisplayAlerts = false;
                    objAssemblyDocument = objDocuments.Open(asmFileFULLPath);

                    if (objAssemblyDocument.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        objAssemblyDocument.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + asmFileFULLPath, logFilePath);
                            return ;
                        }
                    }

                    Utility.Log("DEBUG - InputFile is Opened : " + asmFileFULLPath, logFilePath);
                }
                else
                {
                    Utility.Log("Assembly File Does not Exist " + asmFileFULLPath, logFilePath);
                    return ;
                }

                String outputFile = Path.ChangeExtension(InputFile, "." + Format);
                outputFile = Path.GetFileName(outputFile);
                Utility.Log("DEBUG - outputFile : " + outputFile, logFilePath);
                String outPutFileFULLPath = Path.Combine(OutputFolder, outputFile);

                if (objAssemblyDocument != null)
                {
                    try
                    {

                        //objDraftDocument.SaveAs(outputFile, false, OutPutFileFormat, false, false);
                        if (System.IO.File.Exists(outPutFileFULLPath) == true)
                        {
                            Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath, logFilePath);
                            System.IO.File.Delete(outPutFileFULLPath);
                        }
                        Utility.Log("DEBUG - SaveAs: " + outPutFileFULLPath, logFilePath);
                        objAssemblyDocument.SaveAs(outPutFileFULLPath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath, logFilePath);
                        Utility.Log("Unable to Save " + Format + " for " + itemIDTempStorage, Program.failureFilePath);
                        Utility.Log(ex.Message, logFilePath);
                        return ;
                    }
                    if (File.Exists(outPutFileFULLPath) == false)
                    {
                        Utility.Log("DEBUG - SaveAs Failed : " + outPutFileFULLPath, logFilePath);
                        return ;
                    }
                }

                try
                {
                    if (objAssemblyDocument.ReadOnly == false)
                    {
                        Utility.Log("DEBUG - Save", logFilePath);
                        objAssemblyDocument.Save();
                    }

                }
                catch (Exception ex)
                {
                    Utility.Log("DEBUG - Unable to Save the Document : " + ex.Message, logFilePath);
                    Utility.Log("Unable to Save the assembly file of " + itemIDTempStorage, Program.failureFilePath);
                    return ;
                }
                try
                {
                    Utility.Log("DEBUG - Close", logFilePath);
                    objAssemblyDocument.Close(true);
                   

                }
                catch (Exception ex)
                {
                    Utility.Log("DEBUG - Unable to Close the Document : " + ex.Message, logFilePath);
                    Utility.Log("Unable to Close the assembly file of : " + itemIDTempStorage, Program.failureFilePath);
                    return ;
                }

                Marshal.ReleaseComObject(objAssemblyDocument);
                //Marshal.FinalReleaseComObject(objApp.Documents);
                objAssemblyDocument = null;

                try
                {
                    Utility.Log("DEBUG - Quit", logFilePath);
                    objApp.Quit();
                    OleMessageFilter.Unregister();
                }
                catch (Exception ex)
                {
                    Utility.Log("DEBUG - Quit : " + ex.Message, logFilePath);
                    return;
                }

                Utility.Log("SOL_SE_TRANSLATE: Assembly Save As Completed @ " + System.DateTime.Now.ToString(), logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("DEBUG - Unable to Open : " + asmFileFULLPath, logFilePath);
                Utility.Log("Unable to Open assembly file of " + itemIDTempStorage, Program.failureFilePath);
                Utility.Log(ex.Message, logFilePath);
                return ;
            }

            return ;
        }

        public void SavePsmAsSTATThread(String InputFile, String Format, String logFilePath, String OutputFolder, string itemID)
        {
            itemIDTempStorage = itemID;
            InvokeSTAThread<String, String, String, String>(SavePsmAs, InputFile, Format, logFilePath, OutputFolder);
        }

        [STAThread]
        public void SavePsmAs(String InputFile, String Format, String logFilePath, String OutputFolder)
        {

            String StageDir = Path.GetDirectoryName(InputFile);
            SolidEdge.Framework.Interop.Documents objDocuments = null;
            
            SolidEdge.Part.Interop.SheetMetalDocument shDocument = null;


            String partFileFULLPath = InputFile; // Path.Combine(StageDir, InputFile);
            try
            {
                OleMessageFilter.Register();
                SolidEdge.Framework.Interop.Application objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Start();
                if (objApp == null) return;

                objDocuments = objApp.Documents;
                if (objDocuments == null) return;

                if (System.IO.File.Exists(partFileFULLPath) == true)
                {
                    objApp.DisplayAlerts = false;
                    //objPartDocument = objDocuments.Open(partFileFULLPath);
                    shDocument = objDocuments.Open(partFileFULLPath);
                    if (shDocument.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        shDocument.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + partFileFULLPath, logFilePath);
                            return ;
                        }
                    }

                    Utility.Log("DEBUG - InputFile is Opened : " + partFileFULLPath, logFilePath);
                }
                else
                {
                    Utility.Log("Part File Does not Exist " + partFileFULLPath, logFilePath);
                    return ;
                }

                String outputFile = Path.ChangeExtension(InputFile, "." + Format);
                outputFile = Path.GetFileName(outputFile);
                Utility.Log("DEBUG - outputFile : " + outputFile, logFilePath);
                String outPutFileFULLPath = Path.Combine(OutputFolder, outputFile);

                // ============================================================= DXF
                if (Format.Equals("DXF", StringComparison.OrdinalIgnoreCase) == true)
                {

                    Utility.Log("DEBUG - Save As DXF Starting.. ", logFilePath);

                    SolidEdgePart.FlatPatternModels objFlatPats = null;
                    //SolidEdgePart.FlatPatternModel objFlatPat = null;
                    SolidEdgePart.Models objModels = null;
                    SolidEdgeGeometry.Face objFace = null;
                    SolidEdgeGeometry.Edge objEdge = null;
                    SolidEdgeGeometry.Vertex objVertex = null;

                    if (shDocument != null)
                    {
                        try
                        {
                            objFlatPats = (SolidEdgePart.FlatPatternModels)shDocument.FlatPatternModels;

                            if (objFlatPats != null && objFlatPats.Count > 0)
                            {
                                if (System.IO.File.Exists(outPutFileFULLPath) == true)
                                {
                                    Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath, logFilePath);
                                    System.IO.File.Delete(outPutFileFULLPath);
                                }

                                objModels = (SolidEdgePart.Models)shDocument.Models;
                                if (objModels != null)
                                {
                                    objModels.SaveAsFlatDXFEx(outPutFileFULLPath, objFace, objEdge, objVertex, true);
                                }

                            }
                            else
                            {
                                Utility.Log("Flat Pattern could not be found", logFilePath);
                            }

                            Marshal.ReleaseComObject(objFlatPats);
                            objFlatPats = null;

                        }
                        catch (Exception ex)
                        {
                            Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath, logFilePath);
                            Utility.Log("Unable to Save " + Format + " for " + itemIDTempStorage, Program.failureFilePath);
                            Utility.Log(ex.Message, logFilePath);
                            return;
                        }

                        Utility.Log("DEBUG - Save As DXF Ending.. ", logFilePath);
                    }
                }
                // ============================================================= DXF


                // ============================================================= STP
                if (Format.Equals("STP", StringComparison.OrdinalIgnoreCase) == true)
                {
                    if (shDocument != null)
                    {
                        Utility.Log("DEBUG - Save As STP starting.. ", logFilePath);
                        try
                        {

                            //objDraftDocument.SaveAs(outputFile, false, OutPutFileFormat, false, false);
                            if (System.IO.File.Exists(outPutFileFULLPath) == true)
                            {
                                Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath, logFilePath);
                                System.IO.File.Delete(outPutFileFULLPath);
                            }
                            Utility.Log("DEBUG - SaveAs: " + outPutFileFULLPath, logFilePath);
                            shDocument.SaveAs(outPutFileFULLPath);
                        }
                        catch (Exception ex)
                        {
                            Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath, logFilePath);
                            Utility.Log("Unable to Save " + Format + " for " + itemIDTempStorage, Program.failureFilePath);
                            Utility.Log(ex.Message, logFilePath);
                            return;
                        }

                        Utility.Log("DEBUG - Save As STP Ending.. ", logFilePath);

                    }

                }
                // ============================================================= STP

                if (File.Exists(outPutFileFULLPath) == false)
                    {
                        Utility.Log("DEBUG - SaveAs to "+ Format + " Failed : " + outPutFileFULLPath, logFilePath);
                        //return ;
                    }

                    try
                    {
                        if (shDocument.ReadOnly == false)
                        {
                            Utility.Log("DEBUG - Save", logFilePath);
                            shDocument.Save();
                        }

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Save the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to Save the sheet metal file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }
                    try
                    {
                        Utility.Log("DEBUG - Close", logFilePath);
                        shDocument.Close(true);

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Close the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to close the sheet metal file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }
                 

                    Marshal.ReleaseComObject(shDocument);
                    //Marshal.FinalReleaseComObject(objApp.Documents);
                    shDocument = null;

                try
                {
                    objApp.Quit();

                    OleMessageFilter.Unregister();
                }
                catch (Exception ex)
                {
                    Utility.Log("DEBUG - Quit : " + ex.Message, logFilePath);
                }

            } catch (Exception ex)
            {
                Utility.Log("DEBUG - Unable to Open : " + partFileFULLPath, logFilePath);
                Utility.Log("Unable to open the sheet metal file of " + itemIDTempStorage, Program.failureFilePath);
                Utility.Log(ex.Message, logFilePath);
                return;
            }

            Utility.Log("SOL_SE_TRANSLATE: sheetMetal Save As Completed @ " + System.DateTime.Now.ToString(), logFilePath);

            }

        public void SaveWeldmentAsSTATThread( String InputFile, String Format, String logFilePath, String OutputFolder, string itemID)
        {
            itemIDTempStorage = itemID;
            InvokeSTAThread<String, String, String, String>(SaveWeldmentAs, InputFile, Format, logFilePath, OutputFolder);
        }

        [STAThread]
        public void SaveWeldmentAs( String InputFile, String Format, String logFilePath, String OutputFolder)
        {
            String StageDir = Path.GetDirectoryName(InputFile);

            SolidEdge.Framework.Interop.Documents objDocuments = null;
            
            SolidEdge.Part.Interop.WeldmentDocument objWeldment = null;


            String weldmentFileFULLPath = InputFile; //Path.Combine(StageDir, InputFile);
            try
            {
                OleMessageFilter.Register();

                SolidEdge.Framework.Interop.Application objApp = (SolidEdge.Framework.Interop.Application)SolidEdgeCommunity.SolidEdgeUtils.Start();
                if (objApp == null) return;
                 objDocuments = objApp.Documents;
                 if (objDocuments == null) return;

                if (System.IO.File.Exists(weldmentFileFULLPath) == true)
                {
                    objApp.DisplayAlerts = false;
                    //objPartDocument = objDocuments.Open(partFileFULLPath);
                    objWeldment = objDocuments.Open(weldmentFileFULLPath);
                    if (objWeldment.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        objWeldment.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + weldmentFileFULLPath, logFilePath);
                            return ;
                        }
                    }
                    Utility.Log("DEBUG - InputFile is Opened : " + weldmentFileFULLPath, logFilePath);
                }
                else
                {
                    Utility.Log("Part File Does not Exist " + weldmentFileFULLPath, logFilePath);
                    return ;
                }

                String outputFile = Path.ChangeExtension(InputFile, "." + Format);
                outputFile = Path.GetFileName(outputFile);
                Utility.Log("DEBUG - outputFile : " + outputFile, logFilePath);
                String outPutFileFULLPath = Path.Combine(OutputFolder, outputFile);

                if (objWeldment != null)
                {
                    try
                    {

                        //objDraftDocument.SaveAs(outputFile, false, OutPutFileFormat, false, false);
                        if (System.IO.File.Exists(outPutFileFULLPath) == true)
                        {
                            Utility.Log("DEBUG - outPutFile Exists Already, Deleting: " + outPutFileFULLPath, logFilePath);
                            System.IO.File.Delete(outPutFileFULLPath);
                        }
                        Utility.Log("DEBUG - SaveAs: " + outPutFileFULLPath, logFilePath);                        
                        objWeldment.SaveAs(outPutFileFULLPath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to SaveAs : " + outPutFileFULLPath, logFilePath);
                        Utility.Log("Unable to Save " + Format + " for " + itemIDTempStorage, Program.failureFilePath);
                        Utility.Log(ex.Message, logFilePath);
                        return ;
                    }
                    if (File.Exists(outPutFileFULLPath) == false)
                    {
                        Utility.Log("DEBUG - SaveAs Failed : " + outPutFileFULLPath, logFilePath);
                        return ;
                    }

                    try
                    {
                        if (objWeldment.ReadOnly == false)
                        {
                            Utility.Log("DEBUG - Save", logFilePath);
                            objWeldment.Save();
                        }

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Save the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to Save the weldment file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }
                    try
                    {
                        Utility.Log("DEBUG - Close", logFilePath);
                        objWeldment.Close(true);

                    }
                    catch (Exception ex)
                    {
                        Utility.Log("DEBUG - Unable to Close the Document : " + ex.Message, logFilePath);
                        Utility.Log("Unable to close the weldment file of " + itemIDTempStorage, Program.failureFilePath);
                        return ;
                    }

                    Marshal.ReleaseComObject(objWeldment);
                    //Marshal.FinalReleaseComObject(objApp.Documents);
                    objWeldment = null;
                    
                }
                try
                {
                    objApp.Quit();
                    OleMessageFilter.Unregister();
                }
                catch (Exception ex)
                {
                    Utility.Log("DEBUG - Quit " + ex.Message, logFilePath);
                }

                

                Utility.Log("SOL_SE_TRANSLATE: Weldment Save As Completed @ " + System.DateTime.Now.ToString(), logFilePath);

            }
            catch (Exception ex)
            {
                Utility.Log("DEBUG - Unable to Open : " + weldmentFileFULLPath, logFilePath);
                Utility.Log("Unable to open the weldment file of " + itemIDTempStorage, Program.failureFilePath);
                Utility.Log(ex.Message, logFilePath);
                return;
            }

            return ;
        }     

        
    }

   
}
