using _3DAutomaticUpdate.model;
using SolidEdgeCommunity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate.opInterfaces
{
    class SolidEdgeSetFeatureState : IsolatedTaskProxy
    {
        public static List<String> occurenceList = new List<string>();


        public void SolidEdgeFeatureSyncFromExcelSTAT(String assemblyFileName, List<FeatureLine> featureLineList, String logFilePath)
        {
            InvokeSTAThread<String, List<FeatureLine>, String>(SolidEdgeFeatureSyncFromExcel, assemblyFileName, featureLineList, logFilePath);
        }

        // 22 OCT ASM Feature Processing not there - ASMDOC.EdgeBarFeatures is not there
        [STAThread]
        public void SolidEdgeFeatureSyncFromExcel(String assemblyFileName, List<FeatureLine> featureLineList, String logFilePath)
        {
            occurenceList.Clear();
            if (featureLineList == null || featureLineList.Count == 0)
            {
                Utility.Log("featureLineList: " + featureLineList.Count, logFilePath);
                return;
            }
            Dictionary<String, List<FeatureLine>> featureDictionary = Utility.BuildFeatureDictionary(featureLineList, logFilePath);

            if (featureDictionary == null || featureDictionary.Count == 0)
            {
                Utility.Log("featureDictionary is NULL: ", logFilePath);
                return;
            }



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

                Utility.Log("DEBUG - InputFile is Opened : " + assemblyFileName, logFilePath);

                if (objAssemblyDocument.ReadOnly == true)
                {
                    bool WriteAccess = false;
                    objAssemblyDocument.SeekWriteAccess(out WriteAccess);
                    if (WriteAccess == false)
                    {
                        Utility.ResetAlerts(objApp, false, logFilePath);
                        Utility.Log("Could not get WriteAccess to--" + assemblyFileName, logFilePath);
                        return;
                    }
                }

                if (objAssemblyDocument != null)
                {

                    occurrences = objAssemblyDocument.Occurrences;
                    //Utility.Log("occurrences.Count: " + occurrences.Count, logFilePath);
                    for (int i = 1; i <= occurrences.Count; i++)
                    {
                        occurrence = occurrences.Item(i);
                        if (occurrence == null) continue;
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

                        //Utility.Log("-----------------------------------------", logFilePath);
                        //Utility.Log("occurenceName--" + occurenceName, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        //Utility.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        //Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);



                        SolidEdgePart.PartDocument partDoc = null;
                        SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                        SolidEdgeAssembly.AssemblyDocument assemDoc = null;


                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                            if (partDoc == null) continue;
                            String path = occurrence.OccurrenceFileName;
                            //Utility.Log("path " + path, logFilePath);
                            if (featureDictionary.ContainsKey(occurenceName))
                            {
                                List<FeatureLine> fsList = null;
                                featureDictionary.TryGetValue(occurenceName, out fsList);
                                if (fsList == null || fsList.Count == 0)
                                {
                                    Utility.Log("fsList is NULL,Skipping " + path, logFilePath);
                                    continue;

                                }
                                //Utility.Log("SETFeatureStatus " + path, logFilePath);
                                SETFeatureStatus(occurrence, path, fsList, logFilePath);
                            }

                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                            if (sheetMetalDoc == null) continue;
                            String path = occurrence.OccurrenceFileName;
                            //Utility.Log("path " + path, logFilePath);
                            if (featureDictionary.ContainsKey(occurenceName))
                            {
                                List<FeatureLine> fsList = null;
                                featureDictionary.TryGetValue(occurenceName, out fsList);
                                if (fsList == null || fsList.Count == 0)
                                {
                                    Utility.Log("fsList is NULL,Skipping " + path, logFilePath);
                                    continue;

                                }
                                //Utility.Log("SETFeatureStatus " + path, logFilePath);
                                SETFeatureStatus(occurrence, path, fsList, logFilePath);

                            }


                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                        {
                            // ASM Level Features are not Available.
                            assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            if (assemDoc == null) continue;
                            traverseAssemblyToSetFeatures(assemDoc, featureDictionary, logFilePath);
                        }


                        //Utility.Log("-----------------------------------------", logFilePath);
                    }
                }
                else
                {
                    Utility.ResetAlerts(objApp, false, logFilePath);
                    return;
                }

                SaveAndCloseAssembly(objAssemblyDocument, logFilePath);

                occurenceList.Clear();

                if (objApp != null) objApp.DisplayAlerts = false;

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

        public void traverseAssemblyToSetFeatures(SolidEdgeAssembly.AssemblyDocument assemDoc, Dictionary<String, List<FeatureLine>> featureDictionary, String logFilePath)
        {


            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            if (assemDoc == null)
            {
                Utility.Log("assemDoc is Empty ", logFilePath);
                return;
            }

            occurrences = assemDoc.Occurrences;
            if (occurrences == null)
            {
                Utility.Log("occurrences is Empty " + assemDoc.Name, logFilePath);
                return;
            }
            //Utility.Log("occurrences.Count: " + occurrences.Count, logFilePath);

            for (int i = 1; i <= occurrences.Count; i++)
            {
                occurrence = occurrences.Item(i);
                if (occurrence == null) continue;
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

                Utility.Log("-----------------------------------------", logFilePath);
                Utility.Log("occurenceName--" + occurenceName, logFilePath);
                int occurenceQty = occurrence.Quantity;
                //Utility.Log("occurenceQty--" + occurenceQty, logFilePath);
                String ocurenceFileName = occurrence.OccurrenceFileName;
                Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);



                SolidEdgePart.PartDocument partDoc = null;
                SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;

                try
                {
                    if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                    {

                    }
                }
                catch (Exception ex)
                {
                    Utility.Log("Could not convert occurrence" + assemDoc.Name, logFilePath);
                    continue;
                }


                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                    if (partDoc == null) continue;
                    String path = occurrence.OccurrenceFileName;
                     Utility.Log("path " + path, logFilePath);
                    if (featureDictionary.ContainsKey(occurenceName))
                    {
                        List<FeatureLine> fsList = null;
                        featureDictionary.TryGetValue(occurenceName, out fsList);
                        if (fsList == null || fsList.Count == 0)
                        {
                            Utility.Log("fsList is NULL,Skipping " + path, logFilePath);
                            continue;

                        }
                        //Utility.Log("SETFeatureStatus " + path, logFilePath);
                        SETFeatureStatus(occurrence, path, fsList, logFilePath);
                    }


                    savePart(partDoc, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                    if (sheetMetalDoc == null) continue;
                    String path = occurrence.OccurrenceFileName;
                    Utility.Log("path " + path, logFilePath);
                    if (featureDictionary.ContainsKey(occurenceName))
                    {
                        List<FeatureLine> fsList = null;
                        featureDictionary.TryGetValue(occurenceName, out fsList);
                        if (fsList == null || fsList.Count == 0)
                        {
                            Utility.Log("fsList is NULL,Skipping " + path, logFilePath);
                            continue;

                        }
                        //Utility.Log("SETFeatureStatus " + path, logFilePath);
                        SETFeatureStatus(occurrence, path, fsList, logFilePath);
                    }

                    saveSheet(sheetMetalDoc, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    if (assemDoc1 == null) continue;
                    traverseAssemblyToSetFeatures(assemDoc1, featureDictionary, logFilePath);
                    SaveAndCloseAssembly(assemDoc, logFilePath);
                }


                //Utility.Log("-----------------------------------------", logFilePath);
            }

        }


        // Feature Extract and Read Suppress 
        public void SETFeatureStatus(SolidEdgeAssembly.Occurrence occ, String path, List<FeatureLine> fsList, String logFilePath)
        {
            Utility.Log("SETFeatureStatuspath : " + path, logFilePath);
            SolidEdgePart.PartDocument partDocument = null;
            //SolidEdgeAssembly.AssemblyDocument assemDoc = null;
            SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
            SolidEdgePart.EdgebarFeatures edgebarFeatures = null;

            if (logFilePath == null)
            {
                //Utility.Log("logFilePath does not Exist: ", logFilePath);
                return;
            }

            if (path == null)
            {
                Utility.Log("File does not Exist: ", logFilePath);
                return;
            }

            if (occ == null)
            {
                Utility.Log("occ is NULL " + path, logFilePath);
                return;
            }
            if (occ.OccurrenceDocument == null)
            {
                Utility.Log("OccurrenceDocument is NULL " + path, logFilePath);
                return;

            }


            if (occ.OccurrenceDocument is SolidEdgePart.PartDocument)
            {

                partDocument = (SolidEdgePart.PartDocument)occ.OccurrenceDocument;

                if (partDocument == null)
                {
                    Utility.Log("partDocument is Empty : " + path, logFilePath);
                    return;
                }


                //Utility.Log(partDocument.Name, logFilePath);
                edgebarFeatures = partDocument.DesignEdgebarFeatures;
            }

            if (occ.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
            {
                sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occ.OccurrenceDocument;

                if (sheetMetalDoc == null)
                {
                    Utility.Log("sheetMetalDoc is Empty : " + path, logFilePath);
                    return;
                }
                //Utility.Log(sheetMetalDoc.Name, logFilePath);
                edgebarFeatures = sheetMetalDoc.DesignEdgebarFeatures;
            }

            {

                if (edgebarFeatures == null)
                {
                    Utility.Log("edgebarFeatures is Empty : " + path, logFilePath);
                    return;
                }



                // Interate through the features.
                for (int i = 1; i <= edgebarFeatures.Count; i++)
                {
                    //FeatureLine feature = new FeatureLine();
                    //if (occ.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                    //{
                    //    feature.PartName = sheetMetalDoc.Name;
                    //}
                    //else if (occ.OccurrenceDocument is SolidEdgePart.PartDocument)
                    //{
                    //    feature.PartName = partDocument.Name;
                    //}

                    // Get the EdgebarFeature at current index.
                    object edgebarFeature = edgebarFeatures.Item(i);

                    // Get the managed type.
                    var type = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetType(edgebarFeature);

                    //Utility.Log(type.ToString(), logFilePath);

                    if (type.ToString().Equals("SolidEdgePart.RefPlane") == true)
                    {
                        SolidEdgePart.RefPlane rf = (SolidEdgePart.RefPlane)edgebarFeature;
                        string name = " ";
                        try
                        {
                            name = rf.Name;
                            Utility.Log(rf.Name, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }
                    }
                    if (type.ToString().Equals("SolidEdgePart.Pattern") == true)
                    {
                        SolidEdgePart.Pattern Pattern = (SolidEdgePart.Pattern)edgebarFeature;
                        //Utility.Log(cf.Name, logFilePath);

                        //Boolean isEnabled = null;
                        string name = " ";
                        try
                        {
                            name = Pattern.Name;
                            bool isEnabled = getFeatureEnabledStatus(Pattern.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + Pattern.Name + " SystemName " + Pattern.SystemName + " EdgeBarName: " + Pattern.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                Pattern.Suppress = !isEnabled;

                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }




                    }
                    if (type.ToString().Equals("SolidEdgePart.Etch") == true)
                    {
                        SolidEdgePart.Etch Etch = (SolidEdgePart.Etch)edgebarFeature;
                        //Utility.Log(cf.Name, logFilePath);
                        //bool isEnabled;

                        string name = " ";
                        try
                        {
                            name = Etch.Name;
                            bool isEnabled = getFeatureEnabledStatus(Etch.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + Etch.Name + " SystemName " + Etch.SystemName + " EdgeBarName: " + Etch.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                Etch.Suppress = !isEnabled;

                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }





                    }

                    // 2 OCT
                    if (type.ToString().Equals("SolidEdgePart.Flange") == true)
                    {
                        SolidEdgePart.Flange Flange = (SolidEdgePart.Flange)edgebarFeature;
                        //Utility.Log(cf.Name, logFilePath);
                        //bool isEnabled; // = getFeatureEnabledStatus(Flange.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = Flange.Name;
                            bool isEnabled = getFeatureEnabledStatus(Flange.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + Flange.Name + " SystemName " + Flange.SystemName + " EdgeBarName: " + Flange.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                Flange.Suppress = !isEnabled;

                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }





                    }

                    if (type.ToString().Equals("SolidEdgePart.Chamfer") == true)
                    {
                        SolidEdgePart.Chamfer cf = (SolidEdgePart.Chamfer)edgebarFeature;
                        //Utility.Log(cf.Name, logFilePath);
                        //bool isEnabled; // = getFeatureEnabledStatus(cf.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = cf.Name;
                            bool isEnabled = getFeatureEnabledStatus(cf.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                cf.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.ContourFlange") == true)
                    {
                        SolidEdgePart.ContourFlange cf = (SolidEdgePart.ContourFlange)edgebarFeature;
                        //Utility.Log(cf.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(cf.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = cf.Name;
                            bool isEnabled = getFeatureEnabledStatus(cf.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                cf.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.NormalCutout") == true)
                    {
                        SolidEdgePart.NormalCutout nc = (SolidEdgePart.NormalCutout)edgebarFeature;
                        //Utility.Log(nc.Name, logFilePath);
                        //bool isEnabled ;//= getFeatureEnabledStatus(nc.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = nc.Name;
                            bool isEnabled = getFeatureEnabledStatus(nc.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + nc.Name + " SystemName " + nc.SystemName + " EdgeBarName: " + nc.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                nc.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }





                    }

                    if (type.ToString().Equals("SolidEdgePart.Tab") == true)
                    {
                        SolidEdgePart.Tab tab = (SolidEdgePart.Tab)edgebarFeature;
                        //Utility.Log(tab.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(tab.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = tab.Name;
                            bool isEnabled = getFeatureEnabledStatus(tab.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + tab.Name + " SystemName " + tab.SystemName + " EdgeBarName: " + tab.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                tab.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.MirrorCopy") == true)
                    {
                        SolidEdgePart.MirrorCopy mc = (SolidEdgePart.MirrorCopy)edgebarFeature;
                        //Utility.Log(mc.Name, logFilePath);
                        //bool isEnabled;// = getFeatureEnabledStatus(mc.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = mc.Name;
                            bool isEnabled = getFeatureEnabledStatus(mc.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + mc.Name + " SystemName " + mc.SystemName + " EdgeBarName: " + mc.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                mc.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }




                    }


                    if (type.ToString().Equals("SolidEdgePart.Round") == true)
                    {
                        SolidEdgePart.Round round = (SolidEdgePart.Round)edgebarFeature;
                        //Utility.Log(round.Name, logFilePath);
                        //bool isEnabled;// = getFeatureEnabledStatus(round.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = round.Name;
                            bool isEnabled = getFeatureEnabledStatus(round.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + round.Name + " SystemName " + round.SystemName + " EdgeBarName: " + round.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                round.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }




                    }

                    if (type.ToString().Equals("SolidEdgePart.ExtrudedProtrusion") == true)
                    {
                        SolidEdgePart.ExtrudedProtrusion ep = (SolidEdgePart.ExtrudedProtrusion)edgebarFeature;
                        //Utility.Log(ep.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(ep.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = ep.Name;
                            bool isEnabled = getFeatureEnabledStatus(ep.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + ep.Name + " SystemName " + ep.SystemName + " EdgeBarName: " + ep.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                ep.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }




                    }
                    if (type.ToString().Equals("SolidEdgePart.BreakCorner") == true)
                    {
                        SolidEdgePart.BreakCorner bc = (SolidEdgePart.BreakCorner)edgebarFeature;
                        //Utility.Log(bc.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(bc.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = bc.Name;
                            bool isEnabled = getFeatureEnabledStatus(bc.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + bc.Name + " SystemName " + bc.SystemName + " EdgeBarName: " + bc.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                bc.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }




                    }

                    if (type.ToString().Equals("SolidEdgePart.Hole") == true)
                    {
                        SolidEdgePart.Hole h = (SolidEdgePart.Hole)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled ;//= getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.RevolvedCutout") == true)
                    {
                        SolidEdgePart.RevolvedCutout h = (SolidEdgePart.RevolvedCutout)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled; // = getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.ExtrudedCutout") == true)
                    {
                        SolidEdgePart.ExtrudedCutout h = (SolidEdgePart.ExtrudedCutout)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.RevolvedProtrusion") == true)
                    {
                        SolidEdgePart.RevolvedProtrusion h = (SolidEdgePart.RevolvedProtrusion)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.SweptProtrusion") == true)
                    {
                        SolidEdgePart.SweptProtrusion h = (SolidEdgePart.SweptProtrusion)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled;// = getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.UserDefinedPattern") == true)
                    {
                        SolidEdgePart.UserDefinedPattern h = (SolidEdgePart.UserDefinedPattern)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled ; //= getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                    if (type.ToString().Equals("SolidEdgePart.Thread") == true)
                    {
                        SolidEdgePart.Thread h = (SolidEdgePart.Thread)edgebarFeature;
                        //Utility.Log(h.Name, logFilePath);
                        //bool isEnabled; // = getFeatureEnabledStatus(h.Name, logFilePath, fsList);

                        string name = " ";
                        try
                        {
                            name = h.Name;
                            bool isEnabled = getFeatureEnabledStatus(h.Name, logFilePath, fsList);
                            Utility.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName + " IsEnabled " + isEnabled, logFilePath);
                            try
                            {
                                h.Suppress = !isEnabled;
                            }
                            catch (Exception ex)
                            {
                                Utility.Log(name + " Supress FAILURE: " + ex.Message, logFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utility.Log(ex.Message, logFilePath);
                        }


                    }

                }

                if (occ.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    savePart(partDocument, logFilePath);
                }
                else if (occ.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    saveSheet(sheetMetalDoc, logFilePath);
                }

            }

        }



        private void SaveAndCloseAssembly(SolidEdgeAssembly.AssemblyDocument assemblyDoc, String logFilePath)
        {
            try
            {
                if (assemblyDoc.ReadOnly == false)
                {
                    Utility.Log("SaveAndCloseAssembly:" + assemblyDoc.FullName, logFilePath);
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

        private bool getFeatureEnabledStatus(string featureName, string logFilePath, List<FeatureLine> fsList)
        {
            foreach (FeatureLine fl in fsList)
            {
                if (fl.FeatureName == null || fl.FeatureName == null)
                {
                    continue;
                }
                if (fl.FeatureName.Equals(featureName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    if (fl.IsFeatureEnabled != null && fl.IsFeatureEnabled.Equals("") == false)
                    {
                        if (fl.IsFeatureEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            return true;
                        }
                        else if (fl.IsFeatureEnabled.Equals("N", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }

        private void savePart(SolidEdgePart.PartDocument partDoc, String logFilePath)
        {
            try
            {
                if (partDoc.ReadOnly == false)
                {
                    Utility.Log("savePart:" + partDoc.FullName, logFilePath);
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
                    Utility.Log("saveSheet:" + sheetMetalDoc.FullName, logFilePath);
                    sheetMetalDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
            }
        }
    }
}
