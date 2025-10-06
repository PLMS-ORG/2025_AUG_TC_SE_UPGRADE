using DemoAddInTC.utils;
using SolidEdgeCommunity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DemoAddInTC.controller
{
    class SolidEdgeRedefineLinks
    {
        // 01 - OCT - Replace link Between Draft and part/Assembly/PSM after they are Renamed during CTD.
        // When the Files are Renamed, Then they Loose the Link with their Drafts, It needs to be Reestablished.

        public static void ReplaceLinks(String templateDirectoryToSearchForDrafts, String logFilePath)
        {
            
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);

            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            SolidEdge.RevisionManager.Interop.Document document = null;

            String[] draftFiles = Directory.GetFiles(templateDirectoryToSearchForDrafts, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".dft", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();
            if (draftFiles == null || draftFiles.Length == 0)
            {
                //MessageBox.Show("NO TEMPLATES FOUND");
                return;
            }
            foreach (String dftFile in draftFiles)
            {
                //SolidEdge.RevisionManager.Interop.Application app = SE_SESSION.getRevisionManagerSession();
                // Not Searching Sub Directories, Performance Angle---19/8
                Utlity.Log("dftFile: " + dftFile, logFilePath);
                if (objReviseApp == null)
                {
                    Utlity.Log("ReplaceLinks: Revision Manager Application is NULL ", logFilePath);
                    return;
                }

                try
                {
                    document = objReviseApp.Open(dftFile);
                    if (document == null)
                    {
                        Utlity.Log("ReplaceLinks: " + "Document is NULL", logFilePath);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log("ReplaceLinks: " + ex.Message, logFilePath);
                    return;
                }
                SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

                linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

                if (linkDocuments.Count == 0)
                {
                    Utlity.Log("ReplaceLinks: No Linked Documents in : " + dftFile, logFilePath);
                    return;

                }

                for (int i = 1; i <= linkDocuments.Count; i++)
                {
                    SolidEdge.RevisionManager.Interop.Document linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                    Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);
                   

                }

                
                try
                {
                    document.SaveAllLinks();
                   
                }
                catch (Exception ex)
                {
                    Utlity.Log("SaveAllLinks: " + ex.Message, logFilePath);
                    return;
                }


            }
        }


        public static void ReplaceLinks1(String templateDirectoryToSearchForDrafts, List<String> variablePartsList, String Suffix, String logFilePath)
        {

            String[] draftFiles = Directory.GetFiles(templateDirectoryToSearchForDrafts, "*", SearchOption.AllDirectories)
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".dft", StringComparison.OrdinalIgnoreCase)))
                                         .ToArray();
            if (draftFiles == null || draftFiles.Length == 0)
            {
                //MessageBox.Show("NO TEMPLATES FOUND");
                return;
            }
            foreach (String dftFile in draftFiles)
            {
                // Start a Thread here, Since SolidEdge Functionality Runs Only in STA MODE.
                try
                {
                    Utlity.Log("OpenAndChangeSource: " + System.DateTime.Now.ToString(), logFilePath);
                    Thread myThread = new Thread(() => OpenAndChangeSource(templateDirectoryToSearchForDrafts, dftFile, variablePartsList, Suffix, logFilePath));
                    myThread.SetApartmentState(ApartmentState.STA);
                    myThread.Start();
                    myThread.Join();
                }
                catch (Exception ex)
                {
                    Utlity.Log("SolidEdgeRedefineLinks, OpenAndChangeSource: " + ex.Message, logFilePath);
                    return;
                }
               
            }
        }

        // 15 OCT - NOTE- All Drafts are Copied to Single Folder, All ASM/PAR/PSM are Copied to Single Folder
        private static bool OpenAndChangeSource(String templateDirectoryToSearchForDrafts, string DraftFileFULLPath, List<String> variablePartsList, String Suffix, string logFilePath)
        {
            Utlity.Log("OpenAndChangeSource -  DraftFileFULLPath: " + DraftFileFULLPath, logFilePath);
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
            {
                Utlity.Log("DEBUG -  objApp is NULL : " + DraftFileFULLPath, logFilePath);
                return false;
            }
            
            SolidEdgeDraft.DraftDocument objDraftDocument = null;
            objDocuments = objApp.Documents;           
            
            if (objDocuments == null)
            {
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
                            MessageBox.Show("Could not get WriteAccess to--" + DraftFileFULLPath + "Close and reopen the assembly");
                            return false;
                        }
                    }
                }
                else
                {
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
                SolidEdgeDraft.ModelLinks links = (SolidEdgeDraft.ModelLinks)objDraftDocument.ModelLinks;
                if (links == null)
                {
                    Utlity.ResetAlerts(objApp, true, logFilePath);
                    return false;
                }

                foreach (SolidEdgeDraft.ModelLink link in links)
                {
                    Utlity.Log(link.FileName,logFilePath);
                    String LinkFileName = Path.GetFileName(link.FileName);
                    if (variablePartsList.Contains(LinkFileName) == true)
                    {
                        String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(link.FileName);
                        String extn = System.IO.Path.GetExtension(link.FileName);
                        String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                        String newLinkFileFullPath = Path.Combine(templateDirectoryToSearchForDrafts, newPartFileName);
                        if (System.IO.File.Exists(newLinkFileFullPath) )
                        {
                            Utlity.Log("DEBUG - Draft Relinked To : " + newLinkFileFullPath, logFilePath);
                            try
                            {
                                link.ChangeSource(newLinkFileFullPath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("DEBUG - ChangeSource, Exception : " + ex.Message, logFilePath);
                            }
                        }

                    }
                }
            }
            
            try
            {
                if (objDraftDocument.ReadOnly == false)
                {
                    Utlity.Log("DEBUG - Close & Save", logFilePath);
                    objDraftDocument.Close(true);
                }

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
            Marshal.ReleaseComObject(objDraftDocument);
            //Marshal.FinalReleaseComObject(objApp.Documents);
            objDraftDocument = null;
            if (objApp != null) objApp.DisplayAlerts = true;         

            return true;
        }
    }
}
