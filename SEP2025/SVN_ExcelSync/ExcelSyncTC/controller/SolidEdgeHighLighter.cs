using DemoAddIn.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DemoAddIn.controller
{
    class SolidEdgeHighLighter
    {
        public static List<String> occurenceList = new List<string>();
        public static Dictionary<String, SolidEdgeAssembly.Occurrence> OccurenceObjectDictionary = new Dictionary<string, SolidEdgeAssembly.Occurrence>();


        public static void ClearStore()
        {
            occurenceList.Clear();
            OccurenceObjectDictionary.Clear();
        }
        public static int getOccurenceCount()
        {
            return OccurenceObjectDictionary.Keys.Count();
        }
        public static SolidEdgeAssembly.Occurrence getOccurence(String key)
        {
            SolidEdgeAssembly.Occurrence occurence = null;
            bool returnVal = false;
            if (OccurenceObjectDictionary != null)
            {
                returnVal = OccurenceObjectDictionary.TryGetValue(key, out occurence);
                if (returnVal == false)
                {
                    occurence = null;
                }
            }

            return occurence;
        }
        public static void readOccurences(String logFilePath)
        {
            
            occurenceList.Clear();
            OccurenceObjectDictionary.Clear();

            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;


            


            try
            {
                objDocuments = objApp.Documents;
                
                //objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;
                if (objAssemblyDocument == null)
                {
                    Utlity.Log("DEBUG - InputFile is NOT Opened : ", logFilePath);
                    return;
                }


                if (objAssemblyDocument != null)
                {
                    // This is for Top Assembly Alone
                   // Utlity.Log("AssemDoc.Name : " + objAssemblyDocument.Name, logFilePath);                    
                    occurrences = objAssemblyDocument.Occurrences;
                    
                    for (int i = 1; i <= occurrences.Count; i++)
                    {
                        occurrence = occurrences.Item(i);
                        String occurenceName = occurrence.Name;
                        String[] occArr = occurenceName.Split(':');
                        if (occArr.Length == 2)
                        {
                            occurenceName = occArr[0];
                        }
                        if (occurenceList.Contains(occurenceName) == true)
                        {
                            continue;
                        }
                        else
                        {
                            occurenceList.Add(occurenceName);
                        }

                        //Utlity.Log("-----------------------------------------", logFilePath);
                        Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        //Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        //Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            OccurenceObjectDictionary.Add(occurenceName, occurrence);                            
                            
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            OccurenceObjectDictionary.Add(occurenceName, occurrence);
                        }
                        else if (occurrence.OccurrenceFileName.EndsWith(".asm") == true)
                        {
                            SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;
                            assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            OccurenceObjectDictionary.Add(occurenceName, occurrence);
                            traverseAssembly(assemDoc1, logFilePath);
                        }
                        
                        //Utlity.Log("-----------------------------------------", logFilePath);
                    }
                }
                else
                {
                    return;
                }

            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.Source, logFilePath);
            }            

            // Not Closing Assembly here.
        }


        private static void traverseAssembly(SolidEdgeAssembly.AssemblyDocument assemDoc, String logFilePath)
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
                if (occurenceList.Contains(occurenceName) == true)
                {
                    continue;
                }
                else
                {
                    occurenceList.Add(occurenceName);
                }

                //Utlity.Log("-----------------------------------------", logFilePath);
                Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                int occurenceQty = occurrence.Quantity;
                //Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                String ocurenceFileName = occurrence.OccurrenceFileName;
                //Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    OccurenceObjectDictionary.Add(occurenceName, occurrence);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    OccurenceObjectDictionary.Add(occurenceName, occurrence);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    SolidEdgeAssembly.AssemblyDocument assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    OccurenceObjectDictionary.Add(occurenceName, occurrence);
                    traverseAssembly(assemDoc1, logFilePath);

                    
                }
            }

        }


        public static void HighlightOccurence(String assemFileName, String occurenceName) 
        {

            String AssemblyStageDir = System.IO.Path.GetDirectoryName(assemFileName);
            String LogStageDir = Utlity.CreateLogDirectory();
            //String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(assemFileName) + "_" + "HighLight" + ".txt");      
            String logFilePath = "";

            //if (getOccurenceCount() == 0)
            //{
            //    Utlity.Log("--readOccurences-- " , logFilePath);
            //    readOccurences(logFilePath);                
            //}

            // Connect to a running instance of Solid Edge
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            if (objApp == null)
            {
                Utlity.Log("Solid Edge Application Object is NULL " + occurenceName, logFilePath);
                return;
            }
            SolidEdgeAssembly.AssemblyDocument asmDoc = null;
            try
            {
                asmDoc = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;
            }
            catch (Exception ex)
            {
                Utlity.Log("AssemblyDocument Exception: " + ex .Message , logFilePath);
                return;
            }
            if (asmDoc == null)
            {
                Utlity.Log("AssemblyDocument Object is NULL " + occurenceName, logFilePath);
                return;
            }
            SolidEdgeFramework.HighlightSet hs = asmDoc.HighlightSets.Add();
            if (hs == null)
            {
                Utlity.Log("HighlightSet Object is NULL " + occurenceName, logFilePath);
                return;

            }
            SolidEdgeAssembly.Occurrence occ = null;

            if (OccurenceObjectDictionary != null && OccurenceObjectDictionary.Count !=0) 
            {
                Utlity.Log("Trying to Get Occurence Object from Dictionary: " + occurenceName, logFilePath);
                bool OccurenceSuccess = OccurenceObjectDictionary.TryGetValue(occurenceName, out occ);
                if (OccurenceSuccess == false)
                {
                    occ = null;
                    return;
                }
                Utlity.Log("Got Occurrence Object For: " + occurenceName, logFilePath);
            }

            if (occ != null ) {
                hs.RemoveAll();
                hs.Draw();
                try
                {
                    
                    hs.AddItem(occ);
                }
                catch (Exception ex)
                {
                    Utlity.Log("HighLight: Exception: " + ex.Message, logFilePath);
                    return;
                }
                hs.Draw();
                // WAIT and then Remove HighLight after that.
                System.Threading.Thread.Sleep(2000);

            }
            hs.RemoveAll();
            hs.Draw();
            }
    }
}
