using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidEdge.RevisionManager.Interop;

namespace DemoAddInTC.controller
{
    class SolidEdgeDuplicate
    {
        public static void copyLinkedDocumentsToPublishedFolder2(String folderToPublish, String assemblyFileName, String logFilePath)
        {
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            SolidEdge.RevisionManager.Interop.Document document = null;
            var ListOfInputFiles = new List<string>();
            ListOfInputFiles.Add(assemblyFileName);

            var ListOfInputActions = new List<RevisionManager.RevisionManagerAction>();
            ListOfInputActions.Add(RevisionManager.RevisionManagerAction.CopyAllAction);

            var NewFilePathForAllFiles = folderToPublish;

            String file = System.IO.Path.GetFileName(assemblyFileName);
            String newFileName = System.IO.Path.Combine(folderToPublish, file);
            var ListOfNewFileNames = new List<string>();
            ListOfNewFileNames.Add(newFileName);

            document = objReviseApp.OpenFileInRevisionManager(assemblyFileName);
            

            objReviseApp.SetActionForAllFilesInRevisionManager(RevisionManagerAction.CopyAllAction, folderToPublish);
            try
            {
                objReviseApp.SetActionInRevisionManager(1, (object)ListOfInputFiles, (object)ListOfInputActions, (object)ListOfNewFileNames, NewFilePathForAllFiles);
            }
            catch (Exception ex)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
            }

            try
            {
                objReviseApp.PerformActionInRevisionManager();
            }
            catch (Exception ex)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
            }


            document.Close();
            SE_SESSION.killRevisionManager(logFilePath);
        }


        public static void TraverseLinkedDocuments(SolidEdge.RevisionManager.Interop.Document linkDoc)
        {

        }
    }
}
