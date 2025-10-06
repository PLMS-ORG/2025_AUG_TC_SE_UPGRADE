using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC.controller
{
    class SolidEdgeData1
    {

        public static List<String> occurenceList = new List<string>();
        public static List<Variable> ALLvariablesList = new List<Variable>();
        public static Dictionary<String, List<Variable>> variableDictionary = new Dictionary<string, List<Variable>>
            ();
        public static List<BOMLine> bomLineList = new List<BOMLine>();
        public static List<String> assemblyStructureList = new List<string>();
        public static Dictionary<String, String> ocurrencePathDictionary = new Dictionary<string, string>();
        public static String topLineAssemblyFileName;

        public SolidEdgeData1()
        {

        }

        public static void setAssemblyFileName(String assemblyFileName)
        {
            topLineAssemblyFileName = assemblyFileName;
        }

        public static String getAssemblyFileName()
        {
            return topLineAssemblyFileName;
        }

        public static List<Variable> getVariableDetails()
        {
            return ALLvariablesList;
        }

        private static List<String> getPartNames()
        {

            return occurenceList;
        }

        public static Dictionary<String, List<Variable>> getVariablesDictionaryDetails()
        {
            return variableDictionary;
        }

        public static List<BOMLine> getBomLinesList()
        {
            return bomLineList;
        }

        public static void readVariablesForEachOccurence(String assemblyFileName, String logFilePath)
        {
            variableDictionary.Clear();
            occurenceList.Clear();
            ALLvariablesList.Clear();

            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            try
            {
                objDocuments = objApp.Documents;

                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                Utlity.Log("DEBUG - InputFile is Opened : " + assemblyFileName, logFilePath);


                if (objAssemblyDocument.ReadOnly == true)
                {
                    bool WriteAccess = false;
                    objAssemblyDocument.SeekWriteAccess(out WriteAccess);
                    if (WriteAccess == false)
                    {
                        Utlity.Log("Could not get WriteAccess to--" + assemblyFileName, logFilePath);
                        MessageBox.Show("Could not get WriteAccess to--" + assemblyFileName + "Close and reopen the assembly");
                        return;
                    }
                }

                if (objAssemblyDocument != null)
                {
                    // This is for Top Assembly Alone
                    Utlity.Log("AssemDoc.Name : " + objAssemblyDocument.Name, logFilePath);
                    SolidEdgeFramework.UnitsOfMeasure TopUOM = null;
                    Utlity.Log("TopUOM: ", logFilePath);
                    TopUOM = SolidEdgeUOM.getUOM((SolidEdgeFramework.SolidEdgeDocument)objAssemblyDocument, logFilePath);
                    Utlity.Log("ReadAndFillVariables: ", logFilePath);
                    ReadAndFillVariables(TopUOM, objAssemblyDocument.Name, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath);
                    Utlity.Log("ReadAndFillVariables executed: ", logFilePath);
                    try
                    {
                        occurrences = objAssemblyDocument.Occurrences;
                        Utlity.Log("occurrences.Count: " + occurrences.Count, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("Exception: " + ex.ToString(), logFilePath);
                        Utility.Log("no occurences", logFilePath);
                    }

                    for (int i = 1; i <= occurrences.Count; i++)
                    {
                        occurrence = occurrences.Item(i);
                        if (occurrence == null)
                        {
                            Utlity.Log("Skipping, .... Seems to be Empty Document.", logFilePath);
                            continue;
                        }

                        String occurenceName = occurrence.Name;
                        Utlity.Log("occurenceName: " + occurenceName, logFilePath);
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

                        Utlity.Log("-----------------------------------------", logFilePath);
                        Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                        try
                        {
                            if (occurrence.OccurrenceDocument == null)
                            {
                                Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                            continue;
                        }

                        int occurenceQty = 0;
                        try
                        {
                            occurenceQty = occurrence.Quantity;
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                            continue;
                        }

                        Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = "";
                        try
                        {
                            ocurenceFileName = occurrence.OccurrenceFileName;
                        }
                        catch (Exception ex)
                        {
                            Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                            continue;
                        }
                        Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                        SolidEdgePart.PartDocument partDoc = null;
                        SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                        SolidEdgeAssembly.AssemblyDocument assemDoc = null;
                        SolidEdgeFramework.Variables variables = null;
                        SolidEdgeFramework.UnitsOfMeasure uom = null;

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            if (occurrence.OccurrenceDocument != null)
                            {
                                partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                            }
                            if (partDoc == null) continue;
                            variables = (SolidEdgeFramework.Variables)partDoc.Variables;
                            uom = SolidEdgeUOM.getUOM(occurrence, logFilePath);
                            ReadAndFillVariables(uom, occurenceName, variables, logFilePath);
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            if (occurrence.OccurrenceDocument != null)
                            {
                                sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                            }
                            if (sheetMetalDoc == null) continue;
                            variables = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                            uom = SolidEdgeUOM.getUOM(occurrence, logFilePath);
                            ReadAndFillVariables(uom, occurenceName, variables, logFilePath);
                        }
                        else if (occurrence.OccurrenceFileName.EndsWith(".asm") == true)
                        {
                            if (occurrence.OccurrenceDocument != null)
                            {
                                assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            }

                            if (assemDoc == null) continue;
                            variables = (SolidEdgeFramework.Variables)assemDoc.Variables;
                            uom = SolidEdgeUOM.getUOM(occurrence, logFilePath);
                            traverseAssemblyToReadAndFillVariables(uom, assemDoc, variables, logFilePath);
                        }
                        //variableArr.Clear();
                        Utlity.Log("-----------------------------------------", logFilePath);
                    }
                }
                else
                {
                    return;
                }

            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: readVariablesForEachOccurence " + ex.Message, logFilePath);
                Utlity.Log("Exception: readVariablesForEachOccurence " + ex.Source, logFilePath);
            }

        }


        private static void traverseAssemblyToReadAndFillVariables(SolidEdgeFramework.UnitsOfMeasure UOM, SolidEdgeAssembly.AssemblyDocument assemDoc, SolidEdgeFramework.Variables variables, String logFilePath)
        {
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            if (assemDoc == null)
            {
                Utlity.Log("assemDoc is Empty: " + assemDoc.Name, logFilePath);
                return;
            }
            Utlity.Log("assemDoc.Name: " + assemDoc.Name, logFilePath);
            ReadAndFillVariables(UOM, assemDoc.Name, variables, logFilePath);

            occurrences = assemDoc.Occurrences;
            Utlity.Log("occurrences.Count: " + occurrences.Count, logFilePath);
            for (int i = 1; i <= occurrences.Count; i++)
            {
                occurrence = occurrences.Item(i);
                if (occurrence == null)
                {
                    Utlity.Log("Skipping, .... Seems to be Empty Document.", logFilePath);
                    continue;
                }
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

                Utlity.Log("-----------------------------------------", logFilePath);
                Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                try
                {
                    if (occurrence.OccurrenceDocument == null)
                    {
                        Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                    continue;
                }
                int occurenceQty = 0;
                try
                {
                    occurenceQty = occurrence.Quantity;
                }
                catch (Exception ex)
                {
                    Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                    continue;
                }

                Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                String ocurenceFileName = "";
                try
                {
                    ocurenceFileName = occurrence.OccurrenceFileName;
                }
                catch (Exception ex)
                {
                    Utlity.Log("Skipping, " + occurenceName + ".... Seems to be Empty Document.", logFilePath);
                    continue;
                }
                Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                SolidEdgePart.PartDocument partDoc = null;
                SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;
                SolidEdgeFramework.Variables variables1 = null;
                SolidEdgeFramework.UnitsOfMeasure uom1 = null;

                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    if (occurrence.OccurrenceDocument != null)
                    {
                        partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                    }
                    if (partDoc == null) continue;
                    variables1 = (SolidEdgeFramework.Variables)partDoc.Variables;
                    uom1 = SolidEdgeUOM.getUOM(occurrence, logFilePath);
                    ReadAndFillVariables(uom1, occurenceName, variables1, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    if (occurrence.OccurrenceDocument != null)
                    {
                        sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                    }
                    if (sheetMetalDoc == null) continue;
                    variables1 = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                    uom1 = SolidEdgeUOM.getUOM(occurrence, logFilePath);
                    ReadAndFillVariables(uom1, occurenceName, variables1, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    if (occurrence.OccurrenceDocument != null)
                    {
                        assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    }
                    if (assemDoc1 == null) continue;
                    variables1 = (SolidEdgeFramework.Variables)assemDoc1.Variables;
                    uom1 = SolidEdgeUOM.getUOM(occurrence, logFilePath);
                    traverseAssemblyToReadAndFillVariables(uom1, assemDoc1, variables1, logFilePath);
                    //ReadAndFillVariables(assemDoc1.Name, variables1, logFilePath);
                }
            }

        }

        private static void ReadAndFillVariables(SolidEdgeFramework.UnitsOfMeasure UOM, string occurenceName, SolidEdgeFramework.Variables variables, String logFilePath)
        {
            if (variables == null)
            {
                Utlity.Log("variables is NULL ", logFilePath);
                return;
            }

            SolidEdgeFramework.VariableList variableList = null;
            string pFindCriterium = "*";
            object NamedBy = SolidEdgeConstants.VariableNameBy.seVariableNameByBoth;
            object VarType = SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth;
            object CaseInsensitive = false;



            // Execute query.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
            Utlity.Log("Executed query. variableList.Count " + variableList.Count, logFilePath);
            List<Variable> variableArr = new List<Variable>();
            for (int j = 1; j <= variableList.Count; j++)
            {
                Utility.Log("j" + j, logFilePath);
                var item = variableList.Item(j);
                Utility.Log("Item acquired ", logFilePath);

                // Determine the runtime type of the object.
                var itemType = item.GetType();
                Utility.Log("Item type acquired ", logFilePath);
                var objectType = (SolidEdgeFramework.ObjectType)itemType.InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, item, null);
                Utility.Log("objectType" + objectType.ToString(), logFilePath);

                switch (objectType)
                {
                    case SolidEdgeFramework.ObjectType.igDimension:
                        var dimension = (SolidEdgeFrameworkSupport.Dimension)item;
                        if (dimension != null)
                        {
                            String displayName1 = dimension.DisplayName;
                            /*if ( displayName1.ToLower().Trim().Equals("earthing_hole_left_diameter"))
                            {
                                System.Windows.Forms.MessageBox.Show("Alert"); //25-04-2019
                            }*/
                            //Utlity.Log("dimensionName: " + displayName1, logFilePath);
                            //Utlity.Log(dimension.DimensionType.GetType().ToString(), logFilePath);

                            if (dimension.Expose == 1)
                            {
                                //Utlity.Log("dimensionValue: " + dimension.Value, logFilePath);
                            }
                            //Utlity.Log("dimension.ExposeName: " + dimension.ExposeName, logFilePath);
                            //Utlity.Log("dimension.Formula: " + dimension.Formula, logFilePath);
                            //Utlity.Log("dimension.Comment: " + dimension.GetComment(), logFilePath);

                            try
                            {
                                String dimensionType = SolidEdgeUOM.getType(dimension, logFilePath);
                                String dimensionValue = dimension.Value.ToString();
                                String dimensionValue1 = SolidEdgeUOM.FormatUnit(UOM, dimension, logFilePath);

                                Variable varr = new Variable();

                                varr.name = dimension.DisplayName;
                                //varr.value = dimensionValue1; 1-OCT - LTC Needs Units and Value Separate                  
                                varr.rangeLow = "";
                                varr.rangeCondition = 0;
                                varr.rangeHigh = "";
                                varr.systemName = dimension.SystemName;
                                varr.PartName = occurenceName;
                                varr.Formula = dimension.Formula;
                                varr.AddVarToTemplate = false;
                                varr.DefaultValue = dimension.Value.ToString();
                                varr.variableType = "Dim";
                                varr.UnitType = dimensionType;
                                String Unit = SolidEdgeUOM.StripUnitFromValue(dimensionValue1, logFilePath);
                                varr.unit = Unit;
                                String dimensionValue3 = SolidEdgeUOM.StripValueWithoutUnits(dimensionValue1, logFilePath);
                                varr.value = dimensionValue3;
                                //Utlity.Log("variable.AddVariableToTemplate: " + varr.AddVarToTemplate, logFilePath);
                                ALLvariablesList.Add(varr);
                                variableArr.Add(varr);
                                String dimensionValue2 = SolidEdgeUOM.ParseUnit(UOM, varr, logFilePath);

                                /*Utlity.Log("dimensionDName: " + displayName1 + " dimensionSName: " + dimension.SystemName + " dimensionType: " + dimensionType + " DimensionValue: " + dimensionValue + " DimensionValue1: " + dimensionValue1 + " DimensionValue2: " + dimensionValue2 +
                                " Unit: " + varr.unit + " dimensionValue3: " + dimensionValue3, logFilePath);*/

                                //WriteOutExtraProperties(dimension,logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Exception 443: " + ex.ToString(), logFilePath);
                                continue;
                            }
                        }
                        break;
                    case SolidEdgeFramework.ObjectType.igVariable:
                        var variable = (SolidEdgeFramework.variable)item;
                        if (variable != null)
                        {
                            try
                            {
                                String value = "";
                                variable.GetValue(out value);
                                Utility.Log("value: " + value, logFilePath);
                                String displayName = variable.DisplayName;
                                Utility.Log("displayName: " + displayName, logFilePath);
                                /*if (displayName.ToLower().Trim().Equals("earthing_hole_left_diameter"))
                                {
                                    System.Windows.Forms.MessageBox.Show("Alert"); //25-04-2019
                                }*/
                                String lowValue = "";
                                int Condition;
                                String highValue = "";
                                variable.GetRange(out lowValue, out Condition, out highValue);

                                //Utlity.Log("VariableName: " + displayName, logFilePath);
                                //Utlity.Log(variable.UnitsType.GetType().ToString(), logFilePath);
                                //Utlity.Log("Variablevalue: " + value, logFilePath);                            
                                //Utlity.Log("lowValue: " + lowValue, logFilePath);
                                //Utlity.Log("Condition: " + Condition, logFilePath);
                                //Utlity.Log("highValue: " + highValue, logFilePath);
                                //Utlity.Log("variable.SystemName: " + variable.SystemName, logFilePath);
                                //Utlity.Log("variable.Formula: " + variable.Formula, logFilePath);                            
                                String value1 = SolidEdgeUOM.FormatUnit(UOM, variable, logFilePath);




                                Variable varr = new Variable();

                                varr.name = displayName;
                                // varr.value = value1; - 01 OCT, LTC wants Unit and Value Separate
                                varr.rangeLow = lowValue;
                                varr.rangeCondition = Condition;
                                varr.rangeHigh = highValue;
                                varr.systemName = variable.SystemName;
                                varr.PartName = occurenceName;
                                varr.Formula = variable.Formula;
                                varr.AddVarToTemplate = false;
                                varr.DefaultValue = value;
                                varr.variableType = "Var";
                                varr.UnitType = SolidEdgeUOM.getType(variable, logFilePath);
                                //Utlity.Log("variable.AddVariableToTemplate: " + varr.AddVarToTemplate, logFilePath);
                                String Unit = SolidEdgeUOM.StripUnitFromValue(value1, logFilePath);
                                varr.unit = Unit;

                                String value3 = SolidEdgeUOM.StripValueWithoutUnits(value1, logFilePath);
                                varr.value = value3;
                                ALLvariablesList.Add(varr);
                                variableArr.Add(varr);


                                String value2 = SolidEdgeUOM.ParseUnit(UOM, varr, logFilePath);
                                /*Utlity.Log("VariableDName: " + displayName + " VariableSName: " + variable.SystemName + " variableType: " + SolidEdgeUOM.getType(variable, logFilePath) + " variableValue1: " + value + " variableValue2: " + value1 + " variableValue3: " + value2 +
                                " Unit: " + varr.unit + " variableValue4: " + value3, logFilePath);*/

                                //WriteOutExtraProperties(variable,logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Exception 510: " + ex.ToString(), logFilePath);
                                continue;
                            }

                        }
                        break;
                }
            }

            try
            {
                variableDictionary.Add(occurenceName, variableArr);
                Utlity.Log(occurenceName + ":::::" + variableArr.Count.ToString(), logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Error in adding to variabledictionary " + ex.ToString(), logFilePath);
            }


        }



        public static void traverseAssembly(String assemblyFileName, String logFilePath)
        {
            bomLineList.Clear();
            // 19/8 - Purpose - To find repetitive BomLines & Not Add into bomLineList (Object Store)
            // If Added Again into Object Store, Issues Arise in TreeView in the UI.            
            List<String> BomOccurenceList = new List<string>();
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
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

            BOMLine blTop = new BOMLine();
            blTop.FullName = document.FullName;
            blTop.AbsolutePath = document.AbsolutePath;
            blTop.DocNum = document.DocNum;
            blTop.Revision = document.Revision;
            blTop.Status = document.Status.ToString();
            blTop.Status = getStatus(document);
            blTop.level = "0";
            String fileName = System.IO.Path.GetFileName(document.FullName);
            if (BomOccurenceList.Contains(fileName) == false)
            {
                bomLineList.Add(blTop);
                BomOccurenceList.Add(fileName);
            }


            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + assemblyFileName, logFilePath);
                return;

            }
            int level = 1;
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];
                if (linkDocument.FullName.EndsWith(".xlsx") == true)
                {
                    Utlity.Log("Skipping: " + linkDocument.FullName, logFilePath);
                    continue;
                }

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);
                //Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                //Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                //Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                //Utlity.Log("Status: " + linkDocument.Status, logFilePath);
                BOMLine bl = new BOMLine();
                bl.FullName = linkDocument.FullName;
                bl.AbsolutePath = linkDocument.AbsolutePath;
                bl.DocNum = linkDocument.DocNum;
                bl.Revision = linkDocument.Revision;
                bl.Status = getStatus(linkDocument);
                bl.level = level.ToString();
                // 19 -Aug - Modified to Avoid Repetition

                fileName = System.IO.Path.GetFileName(linkDocument.FullName);
                if (BomOccurenceList.Contains(fileName) == false)
                {
                    bomLineList.Add(bl);
                    BomOccurenceList.Add(fileName);
                }


                //bomLineList.Add(bl);
                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument, logFilePath, level, BomOccurenceList);

                }
            }

            SE_SESSION.killRevisionManager(logFilePath);

        }

        private static void traverseLinkDocuments(SolidEdge.RevisionManager.Interop.Document linkDocument, String logFilePath, int level, List<String> BomOccurenceList)
        {
            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + linkDocument.FullName, logFilePath);
                return;
            }
            level = level + 1;
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);
                //Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                //Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                //Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                //Utlity.Log("Status: " + linkDocument.Status, logFilePath);
                BOMLine bl = new BOMLine();
                bl.FullName = linkDocument.FullName;
                bl.AbsolutePath = linkDocument.AbsolutePath;
                bl.DocNum = linkDocument.DocNum;
                bl.Revision = linkDocument.Revision;
                bl.Status = getStatus(linkDocument);
                bl.level = level.ToString();

                // 19 -Aug - Modified to Avoid Repetition

                String fileName = System.IO.Path.GetFileName(linkDocument.FullName);
                if (BomOccurenceList.Contains(fileName) == false)
                {
                    bomLineList.Add(bl);
                    BomOccurenceList.Add(fileName);
                }


                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseLinkDocuments(linkDocument, logFilePath, level, BomOccurenceList);

                }
            }



        }


        private static String getStatus(SolidEdge.RevisionManager.Interop.Document document)
        {
            String status = "";
            switch (document.Status)
            {
                case SolidEdge.RevisionManager.Interop.DocumentStatus.igStatusAvailable:
                    status = "Available";
                    break;
                case SolidEdge.RevisionManager.Interop.DocumentStatus.igStatusBaselined:
                    status = "Baselined";
                    break;
                case SolidEdge.RevisionManager.Interop.DocumentStatus.igStatusInReview:
                    status = "In Review";
                    break;
                case SolidEdge.RevisionManager.Interop.DocumentStatus.igStatusInWork:
                    status = "In Work";
                    break;
                case SolidEdge.RevisionManager.Interop.DocumentStatus.igStatusObsolete:
                    status = "Obsolete";
                    break;
                case SolidEdge.RevisionManager.Interop.DocumentStatus.igStatusReleased:
                    status = "Released";
                    break;
            }
            return status;
        }

        //public static void updateLinkedTemplate2(String assemblyFileName, String logFilePath)
        //{
        //    SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
        //    SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();

        //    String stageDir = System.IO.Path.GetDirectoryName(assemblyFileName);
        //    String xlFile = System.IO.Path.Combine(stageDir, System.IO.Path.GetFileNameWithoutExtension(assemblyFileName) + ".xlsx");

        //    SolidEdge.RevisionManager.Interop.Document document = null;
        //    var ListOfInputFiles = new List<string>();
        //    ListOfInputFiles.Add(assemblyFileName);

        //    var ListOfInputActions = new List<RevisionManager.RevisionManagerAction>();
        //    ListOfInputActions.Add(RevisionManager.RevisionManagerAction.ReplaceAction);

        //    var NewFilePathForAllFiles = stageDir;

        //    String file = System.IO.Path.GetFileName(assemblyFileName);
        //    String newFileName = System.IO.Path.Combine(stageDir, file);
        //    var ListOfNewFileNames = new List<string>();
        //    ListOfNewFileNames.Add(newFileName);

        //    document = objReviseApp.OpenFileInRevisionManager(assemblyFileName);


        //    //objReviseApp.SetActionForAllFilesInRevisionManager(RevisionManagerAction.ReplaceAction, stageDir);
        //    try
        //    {
        //        objReviseApp.SetActionInRevisionManager(1, (object)ListOfInputFiles, (object)ListOfInputActions, (object)ListOfNewFileNames, NewFilePathForAllFiles);

        //    }
        //    catch (Exception ex)
        //    {
        //        Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
        //    }

        //    try
        //    {
        //        objReviseApp.PerformActionInRevisionManager();
        //    }
        //    catch (Exception ex)
        //    {
        //        Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
        //    }


        //    document.Close();
        //    SE_SESSION.killRevisionManager(logFilePath);
        //}



        public static void copyLinkedDocumentsToPublishedFolder2(String folderToPublish, String assemblyFileName, String logFilePath, bool copyExcelTemplate)
        {
            Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + "Copying Linked Documents", logFilePath);
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
            if (objReviseApp == null)
            {
                Utlity.Log("objReviseApp is NULL", logFilePath);
                return;
            }
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
            if (document == null)
            {
                Utlity.Log("document is NULL", logFilePath);
                return;
            }
            Utlity.Log("SetActionForAllFilesInRevisionManager", logFilePath);
            objReviseApp.SetActionForAllFilesInRevisionManager(SolidEdge.RevisionManager.Interop.RevisionManagerAction.CopyAllAction, folderToPublish);
            try
            {
                Utlity.Log("SetActionInRevisionManager", logFilePath);
                objReviseApp.SetActionInRevisionManager(1, (object)ListOfInputFiles, (object)ListOfInputActions, (object)ListOfNewFileNames, NewFilePathForAllFiles);
            }
            catch (Exception ex)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
            }

            try
            {
                Utlity.Log("PerformActionInRevisionManager", logFilePath);
                objReviseApp.PerformActionInRevisionManager();
            }
            catch (Exception ex)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + ex.Message, logFilePath);
            }


            document.Close();
            SE_SESSION.killRevisionManager(logFilePath);

            if (copyExcelTemplate == true)
            {
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + "Copying Excel Template", logFilePath);
                // Copy the XL file to the Published Folder _ Revisit Later        
                String xlFileNameWExtn = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
                String xlFile = xlFileNameWExtn + ".xlsx";
                //System.IO.Path.GetDirectoryName(assemblyFileName);
                String SourceXlFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(assemblyFileName), xlFile);
                String DestXlFilePath = System.IO.Path.Combine(folderToPublish, xlFile);
                if (System.IO.File.Exists(SourceXlFilePath) == true)
                {
                    try
                    {
                        System.IO.File.Copy(SourceXlFilePath, DestXlFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message + ":::" + SourceXlFilePath, logFilePath);
                        Utlity.Log(ex.Message + ":::" + DestXlFilePath, logFilePath);
                    }
                }
                else
                {
                    Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + "Excel Template is not Available", logFilePath);
                }
                Utlity.Log("copyLinkedDocumentsToPublishedFolder2: " + "Copying Excel Template Completed", logFilePath);
            }
            else
            {
                // Copy the XL file to the Published Folder _ Revisit Later        
                String xlFileNameWExtn = System.IO.Path.GetFileNameWithoutExtension(assemblyFileName);
                String xlFile = xlFileNameWExtn + ".xlsx";
                //System.IO.Path.GetDirectoryName(assemblyFileName);
                String SourceXlFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(assemblyFileName), xlFile);
                String DestXlFilePath = System.IO.Path.Combine(folderToPublish, xlFile);
                if (System.IO.File.Exists(DestXlFilePath) == true)
                {
                    try
                    {
                        Utlity.Log("Delete Existing Template in the Destination" + ":::" + SourceXlFilePath, logFilePath);
                        System.IO.File.Delete(DestXlFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log(ex.Message + ":::" + DestXlFilePath, logFilePath);
                    }
                }

            }

        }

        /** -     Duplicate Structure And Drafts        -- **/
        public static void SearchAndcollectdrafts(String assemblyFileName, String destinationtoCopy, String draftSearchDir, String logFilePath)
        {
            assemblyStructureList.Clear();
            SE_SESSION.InitializeSolidEdgeRevisionManagerSession(logFilePath);
            SolidEdge.RevisionManager.Interop.Application objReviseApp = SE_SESSION.getRevisionManagerSession();
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
            // 1 OCT - ADD Assembly in this List, Fix for DFT files for Top Assembly Missing.
            assemblyStructureList.Add(assemblyFileName);

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + assemblyFileName, logFilePath);
                return;

            }
            Utlity.Log("DUPLICATE: Finding Linked Documents For " + assemblyFileName, logFilePath);
            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);
                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseToCollectLinkDocuments(linkDocument, logFilePath);

                }
                if (assemblyStructureList.Contains(linkDocument.FullName) == false)
                {
                    assemblyStructureList.Add(linkDocument.FullName);
                }
            }
            //assemblyStructureList.Clear();
            document.Close();

            Utlity.Log("DUPLICATE: Searching Drafts Now " + assemblyFileName, logFilePath);
            SolidEdge.RevisionManager.Interop.Application app = SE_SESSION.getRevisionManagerSession();
            //// Not Searching Sub Directories, Performance Angle---19/8
            if (app == null)
            {
                Utlity.Log("searchDrafts: Revision Manager Application is NULL ", logFilePath);
                return;
            }
            foreach (String fullName in assemblyStructureList)
            {
                searchDrafts(fullName, draftSearchDir, destinationtoCopy, logFilePath);
            }
            SE_SESSION.killRevisionManager(logFilePath);
            assemblyStructureList.Clear();

            //SE_SESSION.InitializeSolidEdgeRevisionManagerSession1(logFilePath);

            //RevisionManager.Application app = (RevisionManager.Application)SE_SESSION.getRevisionManagerSession1();
            //foreach (String fullName in assemblyStructureList)
            //{
            //    searchDrafts1(app, fullName, draftSearchDir, destinationtoCopy, logFilePath);

            //}
            //SE_SESSION.killRevisionManager1(logFilePath);
            //assemblyStructureList.Clear();


        }

        private static void traverseToCollectLinkDocuments(SolidEdge.RevisionManager.Interop.Document linkDocument, string logFilePath)
        {

            SolidEdge.RevisionManager.Interop.LinkedDocuments linkDocuments = null;

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)linkDocument.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            if (linkDocuments.Count == 0)
            {
                Utlity.Log("No Linked Documents in : " + linkDocument.FullName, logFilePath);
                return;
            }

            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);

                if (linkDocument.FullName.EndsWith(".asm") == true)
                {
                    traverseToCollectLinkDocuments(linkDocument, logFilePath);
                }
                if (assemblyStructureList.Contains(linkDocument.FullName) == false)
                {
                    assemblyStructureList.Add(linkDocument.FullName);
                }
            }

        }

        private static void searchDrafts(String fileName, String searchDraftDirectory, String draftDestinationDirectory, String logFilePath)
        {
            SolidEdge.RevisionManager.Interop.Application app = SE_SESSION.getRevisionManagerSession();
            // Not Searching Sub Directories, Performance Angle---19/8
            if (app == null)
            {
                Utlity.Log("searchDrafts: Revision Manager Application is NULL " + fileName, logFilePath);
                return;
            }
            app.set_WhereUsedCriteria("*.dft", true, searchDraftDirectory);
            Utlity.Log("searchDrafts: FindWhereUsed Start: " + DateTime.Now.ToString("h:mm:ss tt"), logFilePath);
            //SolidEdge.RevisionManager.Interop.Document document = app.OpenFileInRevisionManager(fileName);
            SolidEdge.RevisionManager.Interop.Document doc = app.FindWhereUsed(fileName);

            Utlity.Log("searchDrafts: FindWhereUsed End: " + DateTime.Now.ToString("h:mm:ss tt"), logFilePath);
            if (doc == null)
            {
                Utlity.Log("No Drafts Identified For: " + fileName, logFilePath);
                return;
            }
            Utlity.Log(doc.FullName, logFilePath);
            String destinationFullName = System.IO.Path.Combine(draftDestinationDirectory, System.IO.Path.GetFileName(doc.FullName));
            if (System.IO.File.Exists(destinationFullName) == false)
            {
                System.IO.File.Copy(doc.FullName, destinationFullName);
            }
            //document.Close();
        }



        private static void searchDrafts1(RevisionManager.Application app, String fileName, String searchDraftDirectory, String draftDestinationDirectory, String logFilePath)
        {
            if (app == null)
            {
                Utlity.Log("No Revision Manager Application : " + fileName, logFilePath);
                return;
            }

            // Not Searching Sub Directories, Performance Angle---19/8
            app.set_WhereUsedCriteria("*.dft", true, searchDraftDirectory);

            object documentsUsedByList = null;
            System.IO.FileInfo f = new System.IO.FileInfo(fileName);
            if (f != null)
            {
                //object outputObj = null;
                try
                {

                    //documentsUsedByList = app.FindWhereUsed(fileName); 
                    app.FindWhereUsedDocuments(f, out documentsUsedByList);

                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message + fileName, logFilePath);
                    return;
                }
                if (documentsUsedByList == null)
                {
                    Utlity.Log("No Drafts Identified For: " + fileName, logFilePath);
                    return;
                }

                List<string> myModel = (List<string>)documentsUsedByList;
                if (myModel == null)
                {
                    Utlity.Log("No Drafts Identified For: " + fileName, logFilePath);
                    return;
                }
                foreach (String fileFullName in myModel)
                {
                    Utlity.Log(fileFullName, logFilePath);
                    String destinationFullName = System.IO.Path.Combine(draftDestinationDirectory, System.IO.Path.GetFileName(fileFullName));
                    if (System.IO.File.Exists(destinationFullName) == false)
                    {
                        System.IO.File.Copy(fileName, destinationFullName);
                    }
                }


            }
            else
            {
                Utlity.Log("Unable to Instantiate FileInfo", logFilePath);
                return;
            }
        }

        // Not Working...
        private static void searchDrafts3(RevisionManager.Application app, String fileName, String searchDraftDirectory, String draftDestinationDirectory, String logFilePath)
        {
            if (app == null)
            {
                Utlity.Log("No Revision Manager Application : " + fileName, logFilePath);
                return;
            }

            // Not Searching Sub Directories, Performance Angle---19/8
            app.set_WhereUsedCriteria("*.dft", true, searchDraftDirectory);

            object documentsUsedByList = null;
            int numberOfDocsFound = 0;
            object ListOfTitles = null;
            object ListOfSub = null;
            object ListofModDtes = null;
            System.IO.FileInfo f = new System.IO.FileInfo(fileName);
            if (f != null)
            {
                //object outputObj = null;
                try
                {

                    app.SearchDocuments(false, searchDraftDirectory, true, out documentsUsedByList, out numberOfDocsFound, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, out ListOfTitles, out ListOfSub, out ListofModDtes);

                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message + fileName, logFilePath);
                    return;
                }
                if (documentsUsedByList == null)
                {
                    Utlity.Log("No Drafts Identified For: " + fileName, logFilePath);
                    return;
                }

                List<string> myModel = (List<string>)documentsUsedByList;
                if (myModel == null)
                {
                    Utlity.Log("No Drafts Identified For: " + fileName, logFilePath);
                    return;
                }
                foreach (String fileFullName in myModel)
                {
                    Utlity.Log(fileFullName, logFilePath);
                    String destinationFullName = System.IO.Path.Combine(draftDestinationDirectory, System.IO.Path.GetFileName(fileFullName));
                    if (System.IO.File.Exists(destinationFullName) == false)
                    {
                        System.IO.File.Copy(fileName, destinationFullName);
                    }
                }


            }
            else
            {
                Utlity.Log("Unable to Instantiate FileInfo", logFilePath);
                return;
            }
        }

        private static void searchDrafts4(String fileName, String searchDraftDirectory, String draftDestinationDirectory, String logFilePath)
        {
            SolidEdge.RevisionManager.Interop.Application app = SE_SESSION.getRevisionManagerSession();
            // Not Searching Sub Directories, Performance Angle---19/8
            if (app == null)
            {
                Utlity.Log("searchDrafts: Revision Manager Application is NULL " + fileName, logFilePath);
                return;
            }
            if (searchDraftDirectory.Contains(";") == true)
            {
                String[] result = searchDraftDirectory.Split(';');
                if (result != null && result.Length > 0)
                {
                    foreach (String DirectoryToSearch in result)
                    {
                        Utlity.Log("searchDrafts: DirectoryToSearch: " + DirectoryToSearch, logFilePath);
                        app.set_WhereUsedCriteria("*.dft", true, DirectoryToSearch);
                    }
                }
            }
            else
            {
                Utlity.Log("searchDrafts: DirectoryToSearch: " + searchDraftDirectory, logFilePath);
                app.set_WhereUsedCriteria("*.dft", true, searchDraftDirectory);
            }

            Utlity.Log("searchDrafts: FindWhereUsed Start: " + DateTime.Now.ToString("h:mm:ss tt"), logFilePath);
            //SolidEdge.RevisionManager.Interop.Document document = app.OpenFileInRevisionManager(fileName);
            SolidEdge.RevisionManager.Interop.Document doc = app.FindWhereUsed(fileName);


            Utlity.Log("searchDrafts: FindWhereUsed End: " + DateTime.Now.ToString("h:mm:ss tt"), logFilePath);
            if (doc == null)
            {
                Utlity.Log("No Drafts Identified For: " + fileName, logFilePath);
                return;
            }
            Utlity.Log(doc.FullName, logFilePath);
            String destinationFullName = System.IO.Path.Combine(draftDestinationDirectory, System.IO.Path.GetFileName(doc.FullName));
            if (System.IO.File.Exists(destinationFullName) == false)
            {
                System.IO.File.Copy(doc.FullName, destinationFullName);
            }
            //document.Close();
        }




    }
}

