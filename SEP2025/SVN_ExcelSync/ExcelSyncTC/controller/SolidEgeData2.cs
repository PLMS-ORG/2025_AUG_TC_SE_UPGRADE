using DemoAddIn.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddIn.controller
{
    class SolidEgeData2
    {
        //Used to Get the Linked Document and its FULLPath during CTD, Full Path Used to Run Create Derivative Method--- IMPORTANT
        private static Dictionary<String, String> AssemblyTraversalDictionary = new Dictionary<String, String>();

        private static List<String> bomLineList = new List<String>();

        //Used to Load Components during TVS --- IMPORTANT
        public static List<String> getBomLinesList()
        {
            return bomLineList;
        }

        public static void traverseAssembly(String assemblyFileName, String logFilePath)
        {
            AssemblyTraversalDictionary.Clear();
            bomLineList.Clear();
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            if (objReviseApp == null)
            {
                Utlity.Log("traverseAssembly: " + "Revision Application Object is NULL", logFilePath);
                return;

            }
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

            String topLineDocument = System.IO.Path.GetFileName(document.FullName);
            if (topLineDocument != null || topLineDocument.Equals("") == false)
            {                
                Utlity.Log(topLineDocument, logFilePath);
                bomLineList.Add(topLineDocument);                
            }

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];
            

            if (linkDocuments == null || linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + assemblyFileName, logFilePath);
                return;

            }
            int level = 1;
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];
                String LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
                Utlity.Log(LinkDocumentName, logFilePath);

                if (bomLineList.Contains(LinkDocumentName) == false)
                {
                    bomLineList.Add(LinkDocumentName);
                }
                if (AssemblyTraversalDictionary.ContainsKey(LinkDocumentName) == false)
                {
                    AssemblyTraversalDictionary.Add(LinkDocumentName, linkDocument.FullName);
                }

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument, logFilePath, level);

                }
            }
            document.Close();
            SE_SESSION.killRevisionManager(logFilePath);

        }

        private static void traverseLinkDocuments(SolidEdge.RevisionManager.Interop.Document linkDocument, String logFilePath, int level)
        {
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];
            String LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
            Utlity.Log(LinkDocumentName, logFilePath);
            if (bomLineList.Contains(LinkDocumentName) == false)
            {
                bomLineList.Add(LinkDocumentName);
            }
            if (AssemblyTraversalDictionary.ContainsKey(LinkDocumentName) == false)
            {
                AssemblyTraversalDictionary.Add(LinkDocumentName, linkDocument.FullName);
            }

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + linkDocument.FullName, logFilePath);
                return;
            }
            level = level + 1;
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];               

                
                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument, logFilePath, level);

                }

                LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
                Utlity.Log(LinkDocumentName, logFilePath);

                if (bomLineList.Contains(LinkDocumentName) == false)
                {
                    bomLineList.Add(LinkDocumentName);
                }
                if (AssemblyTraversalDictionary.ContainsKey(LinkDocumentName) == false)
                {
                    AssemblyTraversalDictionary.Add(LinkDocumentName, linkDocument.FullName);
                }


            }
        }


        // Traverse and Enable/Disable Node in Tree View -- CALLED from MyCustomDialog3 & MyCustomDialog4
        public static void traverseAssembly1(String assemblyFileName, String logFilePath, TreeNode parentNode)
        {
            TreeNode childNode;

            AssemblyTraversalDictionary.Clear();
            bomLineList.Clear();
            
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            if (objReviseApp == null)
            {
                Utlity.Log("traverseAssembly: " + "Revision Application Object is NULL", logFilePath);
                return;

            }
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

            String topLineDocument = System.IO.Path.GetFileName(document.FullName);
            if (topLineDocument != null || topLineDocument.Equals("") == false)
            {
                Utlity.Log(topLineDocument, logFilePath);
                bomLineList.Add(topLineDocument);
            }

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];


            if (linkDocuments == null || linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + assemblyFileName, logFilePath);
                return;

            }
            int level = 1;
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                TreeNode t = new TreeNode();
                
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];
                String LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
                Utlity.Log(LinkDocumentName, logFilePath);

                t.Text = LinkDocumentName;
                t.Name = LinkDocumentName;
                t.Tag = LinkDocumentName;

                if (linkDocument.FullName.EndsWith(".xlsx") == true)
                {
                    Utlity.Log(LinkDocumentName + " SKIPPING", logFilePath);
                    continue;
                }

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    t.ImageIndex = 1;
                    t.SelectedImageIndex = 1;
                }
                else if (linkDocument.FullName.EndsWith(".par") == true || linkDocument.FullName.EndsWith(".psm") == true)
                {
                    t.ImageIndex = 0;
                    t.SelectedImageIndex = 0;

                }

                parentNode.Nodes.Add(t);
                childNode = t;

                if (bomLineList.Contains(LinkDocumentName) == false)
                {
                    bomLineList.Add(LinkDocumentName);
                }
                if (AssemblyTraversalDictionary.ContainsKey(LinkDocumentName) == false)
                {
                    AssemblyTraversalDictionary.Add(LinkDocumentName, linkDocument.FullName);
                }



                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments1(linkDocument, logFilePath, level, childNode);

                }
            }
            document.Close();
            SE_SESSION.killRevisionManager(logFilePath);

        }

        // Traverse and Enable/Disable Node in Tree View -- CALLED from MyCustomDialog3 & MyCustomDialog4
        private static void traverseLinkDocuments1(SolidEdge.RevisionManager.Interop.Document linkDocument, String logFilePath, int level,
            TreeNode tNode)
        {
            TreeNode childNode;

            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];
            String LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
            Utlity.Log(LinkDocumentName, logFilePath);
            if (bomLineList.Contains(LinkDocumentName) == false)
            {
                bomLineList.Add(LinkDocumentName);
            }
            if (AssemblyTraversalDictionary.ContainsKey(LinkDocumentName) == false)
            {
                AssemblyTraversalDictionary.Add(LinkDocumentName, linkDocument.FullName);
            }

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + linkDocument.FullName, logFilePath);
                return;
            }
            level = level + 1;
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
                Utlity.Log(LinkDocumentName, logFilePath);

                TreeNode t = new TreeNode();

                t.Text = LinkDocumentName;
                t.Name = LinkDocumentName;
                t.Tag = LinkDocumentName;

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    t.ImageIndex = 1;
                    t.SelectedImageIndex = 1;
                }
                else if (linkDocument.FullName.EndsWith(".par") == true || linkDocument.FullName.EndsWith(".psm") == true)
                {
                    t.ImageIndex = 0;
                    t.SelectedImageIndex = 0;

                }

                tNode.Nodes.Add(t);
                childNode = t;


                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments1(linkDocument, logFilePath, level, childNode);

                }

                if (bomLineList.Contains(LinkDocumentName) == false)
                {
                    bomLineList.Add(LinkDocumentName);
                }
                if (AssemblyTraversalDictionary.ContainsKey(LinkDocumentName) == false)
                {
                    AssemblyTraversalDictionary.Add(LinkDocumentName, linkDocument.FullName);
                }


            }
        }


        // Trial-01 - Not Working

        public static void CopyAndReplaceSuffixForVariableParts(String folderToPublish, String assemblyFileName, List<String> variableParts,String logFilePath)
        {

            String file = System.IO.Path.GetFileName(assemblyFileName);
            String newFileName = System.IO.Path.Combine(folderToPublish, file);

            Utlity.Log("STEP -1: traverseAssembly: ", logFilePath);
            //traverseAssembly(newFileName, logFilePath);
            
            if (SE_SESSION.getRevisionManagerSession() == null)
            {
                SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            }
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            if (objReviseApp == null)
            {
                Utlity.Log("Revision Application Object is NULL", logFilePath);
                return;
            }
            objReviseApp.DisplayAlerts = 0;
            SolidEdge.RevisionManager.Interop.Document document = null;
            var ListOfInputFiles = new List<string>();
            ListOfInputFiles.Add(assemblyFileName);

            var ListOfInputActions = new List<RevisionManager.RevisionManagerAction>();
            ListOfInputActions.Add(RevisionManager.RevisionManagerAction.RenameAction);            
            var NewFilePathForAllFiles = folderToPublish;
            
            var ListOfNewFileNames = new List<string>();
            ListOfNewFileNames.Add(newFileName);

           

            Utlity.Log("STEP-2: CopyAndReplaceSuffixForVariableParts: ", logFilePath);
            try
            {
                // Opening the Copied Files To Add the Suffix
                document = objReviseApp.OpenFileInRevisionManager(newFileName);
            }
            catch (Exception ex)
            {
                Utlity.Log("OpenFileInRevisionManager: " + ex.Message, logFilePath);
                return;
            }

            if (document == null)
            {
                Utlity.Log("Document Object is NULL", logFilePath);
                return;
            }
            //try
            //{
            //    objReviseApp.SetActionForAllFilesInRevisionManager(SolidEdge.RevisionManager.Interop.RevisionManagerAction.RenameAllAction, folderToPublish);
            //}
            //catch (Exception ex)
            //{
            //    Utlity.Log("SetActionForAllFilesInRevisionManager: " + ex.Message, logFilePath);
            //    return;
            //}

            foreach (String part in AssemblyTraversalDictionary.Keys)
            {

                String partFullPath = "";
                Utlity.Log(part, logFilePath);
                bool Success = AssemblyTraversalDictionary.TryGetValue(part, out partFullPath);
                if (Success == true)
                {
                    if (partFullPath == null || partFullPath.Equals("") == true)
                    {
                        Utlity.Log("partFullPath is Empty", logFilePath);
                        continue;
                    }
                    Utlity.Log(partFullPath, logFilePath);
                    if (System.IO.File.Exists(partFullPath) == false)
                    {
                        Utlity.Log(partFullPath + " is Missing", logFilePath);
                        continue;
                    }
                    ListOfInputFiles.Add(partFullPath);
                    ListOfInputActions.Add(RevisionManager.RevisionManagerAction.RenameAction);
                    if (variableParts.Contains(part) == true)
                    {
                        String extn = System.IO.Path.GetExtension(partFullPath);
                        String oldPartName = System.IO.Path.GetFileNameWithoutExtension(partFullPath);
                        // ADD Suffix.
                        String newPartName = oldPartName + "_SUFFIX";
                        String newpartFullPath = System.IO.Path.Combine(folderToPublish, newPartName + extn);
                        ListOfNewFileNames.Add(newpartFullPath);
                        Utlity.Log(newpartFullPath, logFilePath);
                    }
                    else
                    {
                        ListOfNewFileNames.Add(partFullPath);
                        Utlity.Log(partFullPath, logFilePath);

                    }
                    // Sleep for 1 millisecond to avoid filename collision. Only relevant for this example.
                    System.Threading.Thread.Sleep(1);
                }
            }            
            int ret = 0;
            try
            {
                ret = objReviseApp.SetActionInRevisionManager(ListOfInputFiles.Count, ListOfInputFiles.ToArray(), ListOfInputActions.ToArray(), ListOfNewFileNames.ToArray(), NewFilePathForAllFiles);
            }
            catch (Exception ex)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
            }
            Utlity.Log("Ret Code:" + ret.ToString(), logFilePath);
            try
            {
                ret = objReviseApp.PerformActionInRevisionManager();
               
            }
            catch (Exception ex)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
            }
            Utlity.Log("Ret Code:" + ret.ToString(), logFilePath);

            document.Close();
            SE_SESSION.killRevisionManager(logFilePath);
            

        }

        // Trial-02 - Not Working
        public static void CopyAndReplaceSuffixForVariableParts2(String folderToPublish, String assemblyFileName, List<String> variableParts, String logFilePath)
        {
            // For testing purposes, change the path to the .asm.
            var assemblyPath = assemblyFileName;

            // Start Revision Manager.
            var application = new RevisionManager.Application();

            // Open the assembly.
            var assemblyDocument = (RevisionManager.Document)application.OpenFileInRevisionManager(assemblyPath);

            // Get the linked documents.
            var linkedDocuments = (RevisionManager.LinkedDocuments)assemblyDocument.LinkedDocuments;

            // Allocate input arrays.
            var ListOfInputFileNames = new List<string>();
            var ListOfNewFileNames = new List<string>();
            var ListOfInputActions = new List<RevisionManager.RevisionManagerAction>();

            // Process each linked document.
            for (int i = 1; i <= linkedDocuments.Count; i++)
            {
                // Get the specified linked document by index.
                var linkedDocument = (RevisionManager.Document)linkedDocuments.Item[i];
               
                Utlity.Log(linkedDocument.FullName, logFilePath);
                if (linkedDocument.FullName.EndsWith(".xlsx") == true || linkedDocument.FullName.EndsWith(".par") == true)
                {
                    continue;
                }
                // Get the current path, folder path, filename and extension to the linked document.
                var linkedDocumentPath = linkedDocument.AbsolutePath;
                var linkedDocumentDirectory = System.IO.Path.GetDirectoryName(linkedDocumentPath);
                var linkedDocumentFilename = System.IO.Path.GetFileName(linkedDocumentPath);
                var linkedDocumentExtension = System.IO.Path.GetExtension(linkedDocumentPath);

                // Generate a new random filename for the linked document.
                var linkedDocumentNewPath = DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm-ss-fff", System.Globalization.CultureInfo.InvariantCulture);
                linkedDocumentNewPath = System.IO.Path.ChangeExtension(linkedDocumentNewPath, linkedDocumentExtension);
                linkedDocumentNewPath = System.IO.Path.Combine(linkedDocumentDirectory, linkedDocumentNewPath);

                // Sleep for 1 millisecond to avoid filename collision. Only relevant for this example.
                System.Threading.Thread.Sleep(1);

                // Populate the arrays.
                Utlity.Log(linkedDocumentPath, logFilePath);
                Utlity.Log(linkedDocumentNewPath, logFilePath);
                ListOfInputFileNames.Add(linkedDocumentPath);
                ListOfNewFileNames.Add(linkedDocumentNewPath);
                ListOfInputActions.Add(RevisionManager.RevisionManagerAction.RenameAction);
                
            }
           
            // Set the action.
            application.SetActionInRevisionManager(ListOfInputFileNames.Count, ListOfInputFileNames.ToArray(), ListOfInputActions.ToArray(), ListOfNewFileNames.ToArray());

            // Perform the action.
            application.PerformActionInRevisionManager();

            // Close the assembly.
            assemblyDocument.Close();

            // Close Revision Manager.
            application.Quit();
        }

        public static void CopyAndReplaceSuffixForVariableParts3(String folderToPublish, String assemblyFileName, List<String> variableParts, String Suffix, String logFilePath)
        {
            String newAssemblyFilePath = System.IO.Path.Combine(folderToPublish, System.IO.Path.GetFileName(assemblyFileName));
            CopyAndReplaceSuffixForVariableParts3(folderToPublish, newAssemblyFilePath, variableParts, Suffix, logFilePath, "RENAME");

            CopyAndReplaceSuffixForVariableParts3(folderToPublish, newAssemblyFilePath, variableParts, Suffix, logFilePath, "REPLACE");
        }

        public static void CopyAndReplaceSuffixForVariableParts3(String folderToPublish, String assemblyFileName, List<String> variableParts,String Suffix, String logFilePath,String Option)
        {
            Utlity.Log(Option, logFilePath);
            
            // For testing purposes, change the path to the .asm.
            var assemblyPath = assemblyFileName;

            // Start Revision Manager.
            var application = new RevisionManager.Application();

            // Open the assembly.
            var assemblyDocument = (SolidEdge.RevisionManager.Interop.Document)application.OpenFileInRevisionManager(assemblyPath);

            // Get the linked documents.
            var linkedDocuments = (RevisionManager.LinkedDocuments)assemblyDocument.LinkedDocuments;

            // Process each linked document.
            for (int i = 1; i <= linkedDocuments.Count; i++)
            {
                // Get the specified linked document by index.
                var linkedDocument = (SolidEdge.RevisionManager.Interop.Document)linkedDocuments.Item[i];
               
                Utlity.Log(linkedDocument.FullName, logFilePath);
                
                if (linkedDocument.FullName.EndsWith(".xlsx") == true)
                {
                    continue;
                }
                if (linkedDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments2(linkedDocument, variableParts, logFilePath, folderToPublish, Suffix,Option);
                }

                String fileName = System.IO.Path.GetFileName(linkedDocument.FullName);
                
                if (variableParts != null && variableParts.Count > 0)
                {
                    if (variableParts.Contains(fileName) == true)
                    {
                        String linkedDocumentPath = linkedDocument.AbsolutePath;
                        String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(linkedDocumentPath);
                        String extn = System.IO.Path.GetExtension(linkedDocumentPath);
                        String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                        String newpartFileNameFullPath = System.IO.Path.Combine(folderToPublish, newPartFileName);
                        if (Option.Equals("RENAME", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            Utlity.Log(newpartFileNameFullPath, logFilePath);
                            linkedDocument.Rename(newpartFileNameFullPath);
                        }
                        else
                        {
                            Utlity.Log(newpartFileNameFullPath, logFilePath);
                            linkedDocument.Replace(newpartFileNameFullPath);
                        }
                    }
                }

               

            }

            // Perform the action.
            application.PerformActionInRevisionManager();

            // Close the assembly.
            assemblyDocument.Close();

            // Close Revision Manager.
            application.Quit();
            
        }

        private static void traverseLinkDocuments2(SolidEdge.RevisionManager.Interop.Document linkDocument,List<String>variableParts, String logFilePath,String folderToPublish,String Suffix,String Option)
        {
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];
            String LinkDocumentName = System.IO.Path.GetFileName(linkDocument.FullName);
            Utlity.Log(LinkDocumentName, logFilePath);

            String fileName = System.IO.Path.GetFileName(linkDocument.FullName);
            //Utlity.Log(fileName, logFilePath);
            if (variableParts != null && variableParts.Count > 0)
            {
                if (variableParts.Contains(fileName) == true)
                {
                    String linkedDocumentPath = linkDocument.AbsolutePath;
                    String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(linkedDocumentPath);
                    String extn = System.IO.Path.GetExtension(linkedDocumentPath);
                    String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                    String newpartFileNameFullPath = System.IO.Path.Combine(folderToPublish, newPartFileName);
                    if (Option.Equals("RENAME", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        Utlity.Log(newpartFileNameFullPath, logFilePath);
                        linkDocument.Rename(newpartFileNameFullPath);
                    }
                    else
                    {
                        Utlity.Log(newpartFileNameFullPath, logFilePath);
                        linkDocument.Replace(newpartFileNameFullPath);
                    }
                   
                }
            }

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + linkDocument.FullName, logFilePath);
                return;
            }
           
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];
                
                if (variableParts != null && variableParts.Count > 0)
                {
                    if (variableParts.Contains(fileName) == true)
                    {
                        String linkedDocumentPath = linkDocument.AbsolutePath;
                        String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(linkedDocumentPath);
                        String extn = System.IO.Path.GetExtension(linkedDocumentPath);
                        String newPartFileName = fileNameWithoutExtn + Suffix + extn;
                        String newpartFileNameFullPath = System.IO.Path.Combine(folderToPublish, newPartFileName);
                        if (Option.Equals("RENAME", StringComparison.OrdinalIgnoreCase) == true)
                        {                            
                            linkDocument.Rename(newpartFileNameFullPath);
                        }
                        else
                        {
                            linkDocument.Replace(newpartFileNameFullPath);
                        }
                    }
                }

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments2(linkDocument,variableParts, logFilePath, folderToPublish,Suffix,Option);

                }
            }
        }
    }
}
