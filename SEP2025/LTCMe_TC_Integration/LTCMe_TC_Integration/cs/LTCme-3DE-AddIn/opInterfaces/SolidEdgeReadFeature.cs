using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using SolidEdgeCommunity.Extensions;
using System.IO;

namespace DemoAddInTC.opInterfaces
{
    class SolidEdgeReadFeature
    {

        public static List<FeatureLine> fl = new List<FeatureLine>();
        public static Dictionary<String, List<FeatureLine>> featureDictionary = new Dictionary<string, List<FeatureLine>>();

        public static List<FeatureLine> getFeatureLines()
        {
            return fl;
        }

        public static Dictionary<String, List<FeatureLine>> getFeatureDictionary()
        {
            return featureDictionary;
        }

        [STAThread]
        public static void readFeatures(String logFilePath, String TVS_CTE_OPTION)
        {
            featureDictionary.Clear();
            fl.Clear();

            SolidEdgeCommunity.OleMessageFilter.Register();

            List<BOMLine> bomLinesList = null;

            if (TVS_CTE_OPTION.Equals("TVS", StringComparison.OrdinalIgnoreCase) == true)
            {
                bomLinesList = SolidEdgeData1.getBomLinesList();
            }
            else if (TVS_CTE_OPTION.Equals("CTE", StringComparison.OrdinalIgnoreCase) == true)
            {
                bomLinesList = ExcelData.getBomLineList();
            }

            if (bomLinesList == null || bomLinesList.Count == 0)
            {
                Utlity.Log("bomLinesList is NULL: ", logFilePath);
                return;
            }


            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            //SolidEdgeFramework.Application objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
            if (objApp == null)
            {
                Utlity.Log("objApp is NULL: ", logFilePath);
                return;

            }

            try
            {
                objDocuments = objApp.Documents;

                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;
                if (objAssemblyDocument == null)
                {
                    Utlity.Log("DEBUG - InputFile is NOT Opened : ", logFilePath);
                    return;
                }
                //readFeaturesAtAssemblyLevel(objAssemblyDocument, logFilePath);

                foreach (BOMLine bl in bomLinesList)
                {
                    String partFileName = bl.FullName;
                    Utility.Log("bl.FullName " + bl.FullName, logFilePath);
                    string directory = Path.GetDirectoryName(objAssemblyDocument.FullName);
                    Utility.Log("Directory Name " + directory, logFilePath);
                    partFileName = Path.Combine(directory, Path.GetFileName(bl.FullName));
                    Utility.Log("Part Name " + partFileName, logFilePath);

                    if (partFileName == null || partFileName.Equals("") == true)
                    {
                        Utlity.Log("partFileName does not Exist: ", logFilePath);
                        continue;
                    }

                    if (System.IO.File.Exists(partFileName) == false)
                    {
                        Utlity.Log("File does not Exist: " + partFileName, logFilePath);
                        continue;
                    }
                    String path = System.IO.Path.GetFileName(partFileName);
                    if (path == null)
                    {
                        Utlity.Log("File does not Exist: ", logFilePath);
                        continue;
                    }

                    if (path.EndsWith(".par") == true)
                    {
                        Utility.Log("Getting occurence of " + path, logFilePath);
                        SolidEdgeAssembly.Occurrence occ = SolidEdgeHighLighter.getOccurence(path);

                        if (occ == null)
                        {
                            Utlity.Log("occ is NULL " + path, logFilePath);
                            continue;
                        }

                        readFeatures(occ, path, logFilePath);
                    }
                    if (path.EndsWith(".psm") == true)
                    {
                        SolidEdgeAssembly.Occurrence occ = SolidEdgeHighLighter.getOccurence(path);

                        if (occ == null)
                        {
                            Utlity.Log("occ is NULL " + path, logFilePath);
                            continue;
                        }

                        readFeatures(occ, path, logFilePath);


                    }
                    // Assembly Files needs to be Read As Well
                    if (path.EndsWith(".asm") == true)
                    {
                        SolidEdgeAssembly.Occurrence occ = SolidEdgeHighLighter.getOccurence(path);

                        if (occ == null)
                        {
                            Utlity.Log("occ is NULL " + path, logFilePath);
                            continue;
                        }

                        //readFeaturesAtAssemblyLevel(occ, path, logFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
                Utlity.Log(ex.StackTrace, logFilePath);
                return;
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
            }

        }

        private static void WriteDimensionDetails(SolidEdgePart.Flange Flange, string logFilePath)
        {
            int NumOfDimensions = 0;
            Array DimensionArray = Array.CreateInstance(typeof(object), 0);

            try
            {
                Flange.GetDimensions(out NumOfDimensions, ref DimensionArray);
            }
            catch (Exception ex)
            {
                Utlity.Log("GetDimensions: " + ex.Message, logFilePath);
                return;
            }
            Utlity.Log("NumOfDimensions: " + NumOfDimensions, logFilePath);
            if (DimensionArray != null)
            {
                foreach (object Dim in DimensionArray)
                {
                    SolidEdgeFrameworkSupport.Dimension dimension = (SolidEdgeFrameworkSupport.Dimension)Dim;
                    if (dimension == null) continue;
                    Utlity.Log("dimensionDName: " + dimension.DisplayName + " dimensionSName: " + dimension.SystemName + " dimensionType: " + dimension.Type + " DimensionValue: " + dimension.Value, logFilePath);
                }
            }
        }

        // Feature Extract and Read Suppress 
        public static void
            readFeatures(SolidEdgeAssembly.Occurrence occ, String path, String logFilePath)
        {
            SolidEdgePart.PartDocument partDocument = null;
            //SolidEdgeAssembly.AssemblyDocument assemDoc = null;
            SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
            List<string> featureTypesNotToProcess = new List<string>();
            bool fnResult = readConfigFile(logFilePath, out featureTypesNotToProcess);

            if (fnResult == false)
            {
                Utlity.Log("Could not identify features not to process", logFilePath);
                return;
            }
            SolidEdgePart.EdgebarFeatures edgebarFeatures = null;

            if (logFilePath == null)
            {
                //Utlity.Log("logFilePath does not Exist: ", logFilePath);
                return;
            }

            if (path == null)
            {
                Utlity.Log("File does not Exist: ", logFilePath);
                return;
            }

            if (occ == null)
            {
                Utlity.Log("occ is NULL " + path, logFilePath);
                return;
            }
            try
            {
                if (occ.OccurrenceDocument == null)
                {
                    Utlity.Log("OccurrenceDocument is NULL " + path, logFilePath);
                    
                    return;

                }
            }
            catch (Exception ex)
            {
                Utility.Log(ex.ToString(), logFilePath);
                Utlity.Log("OccurrenceDocument cannot be fetched. Exception " + path, logFilePath);
                return;
            }
            

            if (occ.OccurrenceDocument is SolidEdgePart.PartDocument)
            {
                partDocument = (SolidEdgePart.PartDocument)occ.OccurrenceDocument;

                if (partDocument == null)
                {
                    Utlity.Log("partDocument is Empty : " + path, logFilePath);
                    return;
                }
                Utlity.Log(partDocument.Name, logFilePath);
                edgebarFeatures = partDocument.DesignEdgebarFeatures;
            }

            if (occ.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
            {
                sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occ.OccurrenceDocument;

                if (sheetMetalDoc == null)
                {
                    Utlity.Log("sheetMetalDoc is Empty : " + path, logFilePath);
                    return;
                }
                Utlity.Log(sheetMetalDoc.Name, logFilePath);
                edgebarFeatures = sheetMetalDoc.DesignEdgebarFeatures;
            }

            {

                if (edgebarFeatures == null)
                {
                    Utlity.Log("edgebarFeatures is Empty : " + path, logFilePath);
                    return;
                }


                List<FeatureLine> featuresInPartList = new List<FeatureLine>();
                featuresInPartList.Clear();
                // Interate through the features.
                for (int i = 1; i <= edgebarFeatures.Count; i++)
                {
                    FeatureLine feature = new FeatureLine();
                    if (occ.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                    {
                        feature.PartName = sheetMetalDoc.Name;
                    }
                    else if (occ.OccurrenceDocument is SolidEdgePart.PartDocument)
                    {
                        feature.PartName = partDocument.Name;
                    }

                    // Get the EdgebarFeature at current index.
                    object edgebarFeature = edgebarFeatures.Item(i);

                    // Get the managed type.
                    var type = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetType(edgebarFeature);
                    Utlity.Log(type.ToString() + "was identified", logFilePath);

                    //Checking if feature type is required
                    if (featureTypesNotToProcess.Contains(type.ToString()) == false)
                    {

                        //26-04-2019//
                        string featureName_New = null, formual_New = null, systemName_New = null, edgeBarName_New = null;
                        bool isSuppressed_New = true;


                        Utlity.Log(featureName_New + " " + formual_New + " " + isSuppressed_New + " " + systemName_New + " " + edgeBarName_New, logFilePath);


                        //featureName_New = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<string>(edgebarFeature, "Name", "nil");
                        feature.FeatureName = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<string>(edgebarFeature, "Name", "nil");


                        //formual_New = "";
                        feature.Formula = "";

                        try
                        {
                            isSuppressed_New = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<Boolean>(edgebarFeature, "Suppress", true);
                            if (isSuppressed_New == true)
                                feature.IsFeatureEnabled = "N";
                            else
                                feature.IsFeatureEnabled = "Y";
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }


                        //systemName_New = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<string>(edgebarFeature, "SystemName", "nil");
                        feature.SystemName = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<string>(edgebarFeature, "SystemName", "nil");


                        //edgeBarName_New = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<string>(edgebarFeature, "EdgebarName", "nil");
                        feature.EdgeBarName = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetPropertyValue<string>(edgebarFeature, "EdgebarName", "nil");

                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }

                    //26-04-2019

                    /*
                    if (type.ToString().Equals("SolidEdgePart.Etch") == true)
                    {

                        SolidEdgePart.Etch etch = (SolidEdgePart.Etch)edgebarFeature;
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName, logFilePath);                        
                        //Utlity.Log(flange.Name, logFilePath);
                        feature.FeatureName = etch.Name;

                        feature.Formula = "";
                        try
                        {
                            if (etch.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }



                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = etch.SystemName;
                        feature.EdgeBarName = etch.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }

                    if (type.ToString().Equals("SolidEdgePart.Pattern") == true)
                    {

                        SolidEdgePart.Pattern pattern = (SolidEdgePart.Pattern)edgebarFeature;
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName, logFilePath);                        
                        //Utlity.Log(flange.Name, logFilePath);
                        feature.FeatureName = pattern.Name;

                        feature.Formula = "";
                        try
                        {
                            if (pattern.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }



                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = pattern.SystemName;
                        feature.EdgeBarName = pattern.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }

                    // 02 - OCT, Flange Missing Issue
                    if (type.ToString().Equals("SolidEdgePart.Flange") == true)
                    {
                        
                        SolidEdgePart.Flange flange = (SolidEdgePart.Flange)edgebarFeature;
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName, logFilePath);                        
                        //Utlity.Log(flange.Name, logFilePath);
                        feature.FeatureName = flange.Name;
                        
                        feature.Formula = "";
                        try
                        {
                            if (flange.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                            
                            

                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = flange.SystemName;
                        feature.EdgeBarName = flange.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }
                    else if (type.ToString().Equals("SolidEdgePart.Chamfer") == true)
                    {
                        
                        SolidEdgePart.Chamfer cf = (SolidEdgePart.Chamfer)edgebarFeature;
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName, logFilePath);                        
                        //Utlity.Log(cf.Name, logFilePath);
                        feature.FeatureName = cf.Name;
                        feature.Formula = "";
                        try
                        {
                            if (cf.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                            
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = cf.SystemName;
                        feature.EdgeBarName = cf.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }else if (type.ToString().Equals("SolidEdgePart.ContourFlange") == true)
                    {
                        SolidEdgePart.ContourFlange cf = (SolidEdgePart.ContourFlange)edgebarFeature;
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + cf.Name + " SystemName " + cf.SystemName + " EdgeBarName: " + cf.EdgebarName, logFilePath);
                        //Utlity.Log(cf.Name, logFilePath);
                        feature.FeatureName = cf.Name;                                               
                        feature.Formula = "";
                        try
                        {
                            if (cf.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = cf.SystemName;
                        feature.EdgeBarName = cf.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }else if (type.ToString().Equals("SolidEdgePart.RevolvedCutout") == true)
                    {
                        SolidEdgePart.RevolvedCutout rc = (SolidEdgePart.RevolvedCutout)edgebarFeature;

                        //Utlity.Log(rc.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + nc.Name + " SystemName " + nc.SystemName + " EdgeBarName: " + nc.EdgebarName, logFilePath);
                        feature.FeatureName = rc.Name;
                        feature.Formula = "";
                        try
                        {
                            if (rc.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }

                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = rc.SystemName;
                        feature.EdgeBarName = rc.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);

                    }
                    

                    else if (type.ToString().Equals("SolidEdgePart.NormalCutout") == true)
                    {
                        SolidEdgePart.NormalCutout nc = (SolidEdgePart.NormalCutout)edgebarFeature;

                        //Utlity.Log(nc.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + nc.Name + " SystemName " + nc.SystemName + " EdgeBarName: " + nc.EdgebarName, logFilePath);
                        feature.FeatureName = nc.Name;
                        feature.Formula = "";
                        try
                        {
                            if (nc.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }

                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = nc.SystemName;
                        feature.EdgeBarName = nc.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);

                    }

                    else if (type.ToString().Equals("SolidEdgePart.Tab") == true)
                    {
                        SolidEdgePart.Tab tab = (SolidEdgePart.Tab)edgebarFeature;
                        //Utlity.Log(tab.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + tab.Name + " SystemName " + tab.SystemName + " EdgeBarName: " + tab.EdgebarName, logFilePath);
                        feature.FeatureName = tab.Name;
                        feature.Formula = "";
                        try {
                            if (tab.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }                        
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = tab.SystemName;
                        feature.EdgeBarName = tab.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                    }

                    else if (type.ToString().Equals("SolidEdgePart.MirrorCopy") == true)
                    {
                        SolidEdgePart.MirrorCopy mc = (SolidEdgePart.MirrorCopy)edgebarFeature;
                        //Utlity.Log(mc.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + mc.Name + " SystemName " + mc.SystemName + " EdgeBarName: " + mc.EdgebarName, logFilePath);
                        feature.FeatureName = mc.Name;
                        feature.Formula = "";
                        try
                        {
                            if (mc.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = mc.SystemName;
                        feature.EdgeBarName = mc.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);

                    }

                    else if (type.ToString().Equals("SolidEdgePart.Round") == true)
                    {
                        SolidEdgePart.Round round = (SolidEdgePart.Round)edgebarFeature;
                        //Utlity.Log(round.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + round.Name + " SystemName " + round.SystemName + " EdgeBarName: " + round.EdgebarName, logFilePath);
                        feature.FeatureName = round.Name;
                        feature.Formula = "";
                        try
                        {
                            if (round.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }

                        feature.SystemName = round.SystemName;
                        feature.EdgeBarName = round.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);

                    }

                    else if (type.ToString().Equals("SolidEdgePart.ExtrudedProtrusion") == true)
                    {
                        
                        SolidEdgePart.ExtrudedProtrusion ep = (SolidEdgePart.ExtrudedProtrusion)edgebarFeature;
                        //Utlity.Log(ep.Name, logFilePath);
                      
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + ep.Name + " SystemName " + ep.SystemName +  " EdgeBarName: " + ep.EdgebarName, logFilePath);
                        feature.FeatureName = ep.Name;
                        feature.Formula = "";
                        try {
                            if (ep.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                        {
                            feature.IsFeatureEnabled = "Y";
                        }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = ep.SystemName;
                        feature.EdgeBarName = ep.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);

                    } else if (type.ToString().Equals("SolidEdgePart.BreakCorner") == true)
                    {
                        SolidEdgePart.BreakCorner bc = (SolidEdgePart.BreakCorner)edgebarFeature;
                        //Utlity.Log(bc.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + bc.Name + " SystemName " + bc.SystemName + " EdgeBarName: " + bc.EdgebarName, logFilePath);
                        feature.FeatureName = bc.Name;
                        feature.Formula = "";
                        try
                        {
                            if (bc.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = bc.SystemName;
                        feature.EdgeBarName = bc.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);

                    }

                    else if (type.ToString().Equals("SolidEdgePart.Hole") == true)
                    {
                        SolidEdgePart.Hole h = (SolidEdgePart.Hole)edgebarFeature;
                        //Utlity.Log(h.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName, logFilePath);
                        feature.FeatureName = h.Name;
                        feature.Formula = "";
                        try
                        {
                            if (h.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = h.SystemName;
                        feature.EdgeBarName = h.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                        //Utlity.Log(h.SystemName, logFilePath);
                        //h.Suppress = true;
                    }
                    else if (type.ToString().Equals("SolidEdgePart.ExtrudedCutout") == true)
                    {
                        SolidEdgePart.ExtrudedCutout h = (SolidEdgePart.ExtrudedCutout)edgebarFeature;
                        //Utlity.Log(h.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName, logFilePath);
                        feature.FeatureName = h.Name;
                        feature.Formula = "";
                        try
                        {
                            if (h.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = h.SystemName;
                        feature.EdgeBarName = h.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                        //Utlity.Log(h.SystemName, logFilePath);
                        //h.Suppress = true;
                    }
                    else if (type.ToString().Equals("SolidEdgePart.RevolvedProtrusion") == true)
                    {
                        SolidEdgePart.RevolvedProtrusion h = (SolidEdgePart.RevolvedProtrusion)edgebarFeature;
                        //Utlity.Log(h.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName, logFilePath);
                        feature.FeatureName = h.Name;
                        feature.Formula = "";
                        try
                        {
                            if (h.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = h.SystemName;
                        feature.EdgeBarName = h.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                        //Utlity.Log(h.SystemName, logFilePath);
                        //h.Suppress = true;
                    }
                    else if (type.ToString().Equals("SolidEdgePart.SweptProtrusion") == true)
                    {
                        SolidEdgePart.SweptProtrusion h = (SolidEdgePart.SweptProtrusion)edgebarFeature;
                        //Utlity.Log(h.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName, logFilePath);
                        feature.FeatureName = h.Name;
                        feature.Formula = "";
                        try
                        {
                            if (h.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = h.SystemName;
                        feature.EdgeBarName = h.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                        //Utlity.Log(h.SystemName, logFilePath);
                        //h.Suppress = true;
                    }
                    else if (type.ToString().Equals("SolidEdgePart.UserDefinedPattern") == true)
                    {
                        SolidEdgePart.UserDefinedPattern h = (SolidEdgePart.UserDefinedPattern)edgebarFeature;
                        //Utlity.Log(h.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName, logFilePath);
                        feature.FeatureName = h.Name;
                        feature.Formula = "";
                        try
                        {
                            if (h.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = h.SystemName;
                        feature.EdgeBarName = h.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                        //Utlity.Log(h.SystemName, logFilePath);
                        //h.Suppress = true;
                    }
                    else if (type.ToString().Equals("SolidEdgePart.Thread") == true)
                    {
                        SolidEdgePart.Thread h = (SolidEdgePart.Thread)edgebarFeature;
                        //Utlity.Log(h.Name, logFilePath);
                        //Utlity.Log("TYPE: " + type.ToString() + " Name: " + h.Name + " SystemName " + h.SystemName + " EdgeBarName: " + h.EdgebarName, logFilePath);
                        feature.FeatureName = h.Name;
                        feature.Formula = "";
                        try
                        {
                            if (h.Suppress == true)
                            {
                                feature.IsFeatureEnabled = "N";
                            }
                            else
                            {
                                feature.IsFeatureEnabled = "Y";
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log(ex.Message, logFilePath);
                            feature.IsFeatureEnabled = "Y";
                        }
                        feature.SystemName = h.SystemName;
                        feature.EdgeBarName = h.EdgebarName;
                        fl.Add(feature);
                        featuresInPartList.Add(feature);
                        //Utlity.Log(h.SystemName, logFilePath);
                        //h.Suppress = true;
                    }
                    else
                    {
                        //Utlity.Log("FEATURE SKIPPED : " + type.ToString(), logFilePath);
                        
                    }*/

                }


                if (occ.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    featureDictionary.Add(partDocument.Name, featuresInPartList);
                }
                else if (occ.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    featureDictionary.Add(sheetMetalDoc.Name, featuresInPartList);
                }
                //Utlity.Log("Feature List Count: " + featuresInPartList.Count, logFilePath);

            }

        }

        private static bool readConfigFile(string logFilePath, out List<string> featureTypesNotToProcess) //26-04-2019
        {
            featureTypesNotToProcess = new List<string>();
            try
            {
                string configFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ltcmeAddinConfigFile.config");
                StreamReader reader = new StreamReader(configFilePath);
                string line = null;


                List<string> featureTypesToProcess = new List<string>();
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Length != 0)
                    {
                        if (line.Substring(0, 1).Equals("#"))
                        {
                            if (line.Trim().Equals("#Features that will be processed"))
                            {
                                while ((line = reader.ReadLine()) != null)
                                {
                                    if (line.Length != 0)
                                    {
                                        if (line.Substring(0, 1).Equals("#") == false)
                                        {
                                            string[] contents = line.Split('~');
                                            featureTypesToProcess.Add(contents[0].Trim());
                                        }
                                        else
                                        {
                                            if (line.Trim().Equals("#Features that will not be processed"))
                                            {
                                                while ((line = reader.ReadLine()) != null)
                                                {
                                                    if (line.Length != 0)
                                                    {
                                                        if (line.Substring(0, 1).Equals("#") == false)
                                                        {
                                                            string[] contents = line.Split('~');
                                                            featureTypesNotToProcess.Add(contents[0].Trim());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }

                /*Utlity.Log("Will be processed", logFilePath);
                foreach (string s in featureTypesToProcess)
                    Utlity.Log(s, logFilePath);
                Utlity.Log("\n", logFilePath);*/
                
                /*Utlity.Log("Will not be processed", logFilePath);
                foreach (string s in featureTypesNotToProcess)
                    Utlity.Log(s, logFilePath);*/

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static void WriteDimensionDetails(object edgebarFeature, string logFilePath)
        {
            Type typ = edgebarFeature.GetType(); // This returns type "System.__ComObject

            // Show the name of the feature
            try
            {
                string name = (string)typ.InvokeMember("EdgebarName", BindingFlags.GetProperty, null, edgebarFeature, null);
            }
            catch (Exception ex)
            {
                Utlity.Log("EdgebarName: Could Not Retrieve the Details: " + ex.Message, logFilePath);
                return;
            }

            try
            {
                // When invoking COM-methods with out/ref arguments, you have to explicitely tell this using the
                // ParameterModifier class.

                // Create an array containing the arguments (does not have to be initialized as we are requesting them)
                int numofDims = 0;
                Array dims = Array.CreateInstance(typeof(SolidEdgeFrameworkSupport.Dimension), 0);
                object[] args = new object[2] { numofDims, dims };

                // Initialize a ParameterModifier with the number of arguments
                ParameterModifier argMod = new ParameterModifier(2);

                // Pass the first and second argument by reference
                argMod[0] = true;
                argMod[1] = true;

                // The ParameterModifier must be passed as a single element of an array
                ParameterModifier[] mods = { argMod };

                // Invoke the "GetProfiles" method
                typ.InvokeMember("GetDimensions", BindingFlags.InvokeMethod, null, edgebarFeature, args, mods, null, null);

                // Use the args array to access the returned values
                numofDims = (int)args[0];
                Utlity.Log("    numofDims =" + numofDims, logFilePath);

                if (numofDims > 0)
                {
                    dims = (Array)args[1];
                    SolidEdgeFrameworkSupport.Dimension dim = (SolidEdgeFrameworkSupport.Dimension)dims.GetValue(0);
                    Utlity.Log("       #Name   = {0}" + dim.Name, logFilePath);
                    Utlity.Log("       #SystemName = {0}" + dim.SystemName, logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("   EXC: " + ex.Message, logFilePath);
            }

        }

        private static void readFeaturesAtAssemblyLevel(SolidEdgeAssembly.AssemblyDocument objAssemblyDocument, string logFilePath)
        {
            if (logFilePath == null)
            {
                //Utlity.Log("logFilePath does not Exist: ", logFilePath);
                return;
            }

            SolidEdgeAssembly.AssemblyDrivenPartFeatures AssemblyDrivenPartFeaturesObj = null;


            if (objAssemblyDocument == null)
            {
                Utlity.Log("AssemblyDoc is Empty : ", logFilePath);
                return;
            }
            Utlity.Log(objAssemblyDocument.Name, logFilePath);
            AssemblyDrivenPartFeaturesObj = objAssemblyDocument.AssemblyDrivenPartFeatures;



            if (AssemblyDrivenPartFeaturesObj == null)
            {
                Utlity.Log("AssemblyDrivenPartFeaturesObj is Empty : ", logFilePath);
                return;
            }


            for (int i = 1; i <= AssemblyDrivenPartFeaturesObj.AssemblyDrivenPartFeaturesExtrudedCutouts.Count; i++)
            {
                SolidEdgeAssembly.AssemblyDrivenPartFeaturesExtrudedCutout AssemblyExtrudedCutOut = null;
                AssemblyExtrudedCutOut = AssemblyDrivenPartFeaturesObj.AssemblyDrivenPartFeaturesExtrudedCutouts.Item(i);
                if (AssemblyExtrudedCutOut == null)
                {
                    Utlity.Log("AssemblyFeaturesHole Item is Empty : ", logFilePath);
                    continue;
                }
                FeatureLine AssemblyFeatureLine = new FeatureLine();

                AssemblyFeatureLine.EdgeBarName = AssemblyExtrudedCutOut.Name;
                AssemblyFeatureLine.FeatureName = AssemblyExtrudedCutOut.Name;
                AssemblyFeatureLine.Formula = "";
                try
                {
                    AssemblyExtrudedCutOut.Suppress = false;
                }
                catch (Exception ex)
                {
                    Utlity.Log("AssemblyExtrudedCutOut Exception : " + ex.Message, logFilePath);
                }

                try
                {
                    object Description;
                    SolidEdgeConstants.FeatureStatusConstants Constants = (SolidEdgeConstants.FeatureStatusConstants)AssemblyExtrudedCutOut.get_Status(out Description);
                    AssemblyExtrudedCutOut.get_Status(out Description);
                    if (Constants == SolidEdgeConstants.FeatureStatusConstants.igFeatureSuppressed)
                    {
                        Utlity.Log("AssemblyFeaturesHole Item is igFeatureSuppressed : ", logFilePath);
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log("AssemblyExtrudedCutOut Exception : " + ex.Message, logFilePath);
                    Utlity.Log("AssemblyExtrudedCutOut Exception : " + ex.StackTrace, logFilePath);
                }

            }

            objAssemblyDocument.Save();

            for (int i = 1; i <= AssemblyDrivenPartFeaturesObj.AssemblyDrivenPartFeaturesHoles.Count; i++)
            {
                SolidEdgeAssembly.AssemblyDrivenPartFeaturesHole AssemblyFeatureHole = null;
                AssemblyFeatureHole = AssemblyDrivenPartFeaturesObj.AssemblyDrivenPartFeaturesHoles.Item(i);
                if (AssemblyFeatureHole == null)
                {
                    Utlity.Log("AssemblyFeaturesHole Item is Empty : ", logFilePath);
                    continue;
                }
                FeatureLine AssemblyFeatureLine = new FeatureLine();

                AssemblyFeatureLine.EdgeBarName = AssemblyFeatureHole.Name;
                AssemblyFeatureLine.FeatureName = AssemblyFeatureHole.Name;
                AssemblyFeatureLine.Formula = "";

                //AssemblyFeatureHole.Suppress = true;
                object Description = null;
                SolidEdgeConstants.FeatureStatusConstants Constants = (SolidEdgeConstants.FeatureStatusConstants)AssemblyFeatureHole.get_Status(out Description);
                if (Constants == SolidEdgeConstants.FeatureStatusConstants.igFeatureSuppressed)
                {
                    Utlity.Log("AssemblyFeaturesHole Item is igFeatureSuppressed : ", logFilePath);
                }

            }


        }

    }
}
