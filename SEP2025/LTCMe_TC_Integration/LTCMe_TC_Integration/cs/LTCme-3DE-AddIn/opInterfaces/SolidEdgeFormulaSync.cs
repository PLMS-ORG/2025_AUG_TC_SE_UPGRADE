using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC.opInterfaces
{
    class SolidEdgeFormulaSync
    {
        public static List<String> occurenceList = new List<string>();
        

        public static void SyncSolidEdgeFormula(String assemblyFileName, List<Variable> variablesList, Dictionary<String, bool> partEnablementDictionary, String logFilePath)
        {
            occurenceList.Clear();
            Dictionary<String, List<Variable>> variableDictionaryDetails = Utlity.BuildVariableDictionary(variablesList, logFilePath);

            if (variableDictionaryDetails == null || variableDictionaryDetails.Count == 0)
            {
                Utlity.Log("DEBUG - variableDictionaryDetails is Empty : " + assemblyFileName, logFilePath);
                return;
            }

            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                Utlity.Log("DEBUG - partEnablementDictionary is Empty : " + assemblyFileName, logFilePath);
                return;
            }
            if (variablesList == null || variablesList.Count == 0)
            {
                Utlity.Log("DEBUG - variablesList is Empty : " + assemblyFileName, logFilePath);
                return;
            }
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;
                       
            try
            {
                if (objApp == null)
                {
                    Utlity.Log("DEBUG - objApp is Empty : " + assemblyFileName, logFilePath);
                    return;
                }
                objDocuments = objApp.Documents;
                objApp.DisplayAlerts = false;
                //objApp.Visible = false;
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                Utlity.Log("DEBUG - InputFile is Opened : " + assemblyFileName, logFilePath);

                if (objAssemblyDocument.ReadOnly == true)
                {
                    bool WriteAccess = false;
                    objAssemblyDocument.SeekWriteAccess(out WriteAccess);
                    if (WriteAccess == false)
                    {
                        Utlity.ResetAlerts(objApp, true, logFilePath);
                        Utlity.Log("Could not get WriteAccess to--" + assemblyFileName, logFilePath);
                        MessageBox.Show("Could not get WriteAccess to--" + assemblyFileName + "Close and reopen the assembly");
                        return;
                    }
                }

                if (objAssemblyDocument != null)
                {
                    List<Variable> variablesList1 = null;
                    variableDictionaryDetails.TryGetValue(objAssemblyDocument.Name, out variablesList1);

                    if (variablesList1 == null || variablesList1.Count == 0)
                    {
                        Utlity.Log("VariablesList1 is Empty--" + objAssemblyDocument.Name, logFilePath);

                    }
                    else if (variablesList1.Count != 0)
                    {

                        foreach (Variable varr in variablesList1)
                        {
                            SolidEdgeFramework.UnitsOfMeasure UOM = null;
                            UOM = objAssemblyDocument.UnitsOfMeasure;
                            SetFormula(objAssemblyDocument.Name, varr.systemName, varr.Formula, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath,varr.variableType);

                            try
                            {
                                SetVariableData(UOM,objAssemblyDocument.Name, varr.systemName, varr, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath);                                
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("SetVariableData--" + ex.Message, logFilePath);
                            }

                        }
                    }
                    occurrences = objAssemblyDocument.Occurrences;
                    Utlity.Log("occurrences.Count: " + occurrences.Count, logFilePath);
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
                        String ocurenceFileName = occurrence.OccurrenceFileName;                        


                        SolidEdgePart.PartDocument partDoc = null;
                        SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                        SolidEdgeAssembly.AssemblyDocument assemDoc = null;
                        SolidEdgeFramework.Variables variables = null;

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                            if (partDoc == null)
                            {
                                continue;
                            }
                            variables = (SolidEdgeFramework.Variables)partDoc.Variables;
                            variablesList1 = null;
                            variableDictionaryDetails.TryGetValue(partDoc.Name, out variablesList1);

                            if (variablesList1 == null || variablesList1.Count == 0)
                            {
                                Utlity.Log("variablesList1 is Empty--" + partDoc.Name, logFilePath);

                            }
                            else if (variablesList1.Count != 0)
                            {
                                bool partEnable = false;
                                partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                                if (partEnable == false)
                                {
                                    Utlity.Log("Skipping Sync of " + occurenceName, logFilePath);
                                    continue;
                                }

                                foreach (Variable varr in variablesList1)
                                {
                                    SetFormula(partDoc.Name, varr.systemName, varr.Formula, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath, varr.variableType);

                                    try
                                    {
                                        SetVariableData(partDoc.UnitsOfMeasure,partDoc.Name, varr.systemName, varr, variables, logFilePath);
                                    }
                                    catch (Exception ex)
                                    {
                                        Utlity.Log("SetVariableData--" + ex.Message, logFilePath);
                                    }


                                }
                            }
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                            if (sheetMetalDoc == null)
                            {
                                continue;
                            }

                            variables = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                            variablesList1 = null;
                            variableDictionaryDetails.TryGetValue(sheetMetalDoc.Name, out variablesList1);
                            if (variablesList1 == null || variablesList1.Count == 0)
                            {
                                //Utlity.Log("variablesList1 is Empty--" + sheetMetalDoc.Name, logFilePath);

                            }
                            else if (variablesList1.Count != 0)
                            {
                                bool partEnable = false;
                                partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                                if (partEnable == false)
                                {
                                    Utlity.Log("Skipping Sync of " + occurenceName, logFilePath);
                                    continue;
                                }

                                foreach (Variable varr in variablesList1)
                                {
                                    SetFormula(sheetMetalDoc.Name, varr.systemName, varr.Formula, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, varr.variableType);
                                    try
                                    {
                                        SetVariableData(sheetMetalDoc.UnitsOfMeasure,sheetMetalDoc.Name, varr.systemName, varr, variables, logFilePath);
                                    }
                                    catch (Exception ex)
                                    {
                                        Utlity.Log("SetVariableData--" + ex.Message, logFilePath);
                                    }

                                }

                            }
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                        {
                            assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)assemDoc.Variables;
                            traverseAssemblyToSetVariables(assemDoc, variables, logFilePath, variableDictionaryDetails, partEnablementDictionary);
                        }


                        Utlity.Log("-----------------------------------------", logFilePath);
                    }
                }
                else
                {
                    Utlity.ResetAlerts(objApp, true, logFilePath);
                    return;
                }

            }
            catch (Exception ex)
            {
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.Source, logFilePath);
                Utlity.Log("Exception: " + ex.StackTrace, logFilePath);
                Utlity.ResetAlerts(objApp, true, logFilePath);
            }

            SaveAndCloseAssembly(objAssemblyDocument, logFilePath);

            occurenceList.Clear();
            //partEnablementDictionary.Clear();
            variableDictionaryDetails.Clear();
            if (objApp != null) objApp.DisplayAlerts = true;  

        }

        // Set Formula for the Variable
        private static void SetFormula(string occurenceName, String variableName, String Formula, SolidEdgeFramework.Variables variables, String logFilePath,String variableType)
        {
            if (variables == null)
            {
                Utlity.Log("variables is NULL ", logFilePath);
                return;
            }
            //if (Formula == null || Formula.Equals(""))
            //{
            //    //Utlity.Log("Formula is NULL For: " + variableName, logFilePath);
            //    return;

            //}

            SolidEdgeFramework.VariableList variableList = null;
            string pFindCriterium = "*" + variableName + "*";
            object NamedBy = SolidEdgeConstants.VariableNameBy.seVariableNameByBoth;
            object VarType = SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth;
            object CaseInsensitive = false;

            // Execute query.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
            if (variableList.Count == 0)
            {
                Utlity.Log("variableList is NULL For: " + variableName, logFilePath);
                return;
            }
            var item = variableList.Item(1);

            // Determine the runtime type of the object.
            var itemType = item.GetType();
            if (itemType == null)
            {
                Utlity.Log("itemType is NULL For: " + variableName, logFilePath);
                return;

            }
            var objectType = (SolidEdgeFramework.ObjectType)itemType.InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, item, null);

            switch (objectType)
            {
                case SolidEdgeFramework.ObjectType.igDimension:
                    var dimension = (SolidEdgeFrameworkSupport.Dimension)item;
                    if (dimension != null)
                    {
                        String displayName = dimension.DisplayName;

                        if (Formula.Equals("") == true || Formula == null)
                        {
                            if (dimension.Formula != null || dimension.Formula.Equals("") == false)
                            {
                                //Utlity.Log("Set Formula to Empty: " + dimension.Formula, logFilePath);
                                dimension.Formula = Formula;
                            }
                        }
                        
                        if (Formula.Equals(dimension.Formula, StringComparison.OrdinalIgnoreCase) == false)
                        {
                            dimension.Formula = Formula;
                            //Utlity.Log("DimensionName: " + displayName, logFilePath);
                            //Utlity.Log("Set Formula: " + dimension.Formula, logFilePath);
                        }
                    }
                    break;
                case SolidEdgeFramework.ObjectType.igVariable:
                    var variable = (SolidEdgeFramework.variable)item;
                    if (variable != null)
                    {
                        String displayName = variable.DisplayName;

                        if (Formula.Equals("") == true || Formula == null)
                        {
                            if (variable.Formula != null || variable.Formula.Equals("") == false)
                            {
                                //Utlity.Log("Set Formula to Empty: " + variable.Formula, logFilePath);
                                variable.Formula = Formula;
                            }
                        }

                        if (Formula.Equals(variable.Formula, StringComparison.OrdinalIgnoreCase) == false)
                        {
                            //Utlity.Log("VariableName: " + displayName, logFilePath);
                            variable.Formula = Formula;
                            //Utlity.Log("Set Formula: " + variable.Formula, logFilePath);
                        }
                    }
                    break;
            }

        }


        // Set Formula for the Variable
        private static void SetVariableData(SolidEdgeFramework.UnitsOfMeasure UOM, string occurenceName, String variableName, Variable updatedVar, SolidEdgeFramework.Variables variables, String logFilePath)
        {
            if (variables == null)
            {
                Utlity.Log("variables is NULL ", logFilePath);
                return;
            }
            if (updatedVar == null)
            {
                Utlity.Log("updatedVar is NULL ", logFilePath);
                return;

            }

            // 01- OCT- Changes for Requirement from LTC - Split Unit & Value
            updatedVar.value = SolidEdgeUOM.MergeValueAndUnit(updatedVar.value, updatedVar.unit, logFilePath);

            SolidEdgeFramework.VariableList variableList = null;
            string pFindCriterium = "*" + variableName + "*";
            
            object NamedBy = SolidEdgeConstants.VariableNameBy.seVariableNameByBoth;
            object VarType = SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth;
            object CaseInsensitive = false;
            
            // Execute query.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
            if (variableList == null || variableList.Count == 0)
            {
                
                pFindCriterium = "*" + updatedVar.name + "*";
                variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
                if (variableList == null || variableList.Count == 0)
                {
                    Utlity.Log("variableList is NULL For: " + updatedVar.name, logFilePath);
                    return;
                    //pFindCriterium = "*";
                    //variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
                    //if (variableList == null || variableList.Count == 0)
                    //{
                    //    Utlity.Log("variableList is NULL For: " + updatedVar.name, logFilePath);
                    //    return;
                    //}
                    //else
                    //{
                    //    Utlity.Log("variableList COUNT: " + variableList.Count, logFilePath);
                    //    return;
                    //}
                }
                else
                {
                    Utlity.Log("variableList COUNT: " + variableList.Count, logFilePath);
                }
                
            }

            try
            {
                var item = variableList.Item(1);
                // Determine the runtime type of the object.
                var itemType = item.GetType();
                if (itemType == null)
                {
                    Utlity.Log("itemType is NULL For: " + variableName, logFilePath);
                    return;

                }
                var objectType = (SolidEdgeFramework.ObjectType)itemType.InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, item, null);

                switch (objectType)
                {
                    case SolidEdgeFramework.ObjectType.igDimension:
                        var dimension = (SolidEdgeFrameworkSupport.Dimension)item;
                        if (dimension != null)
                        {
                            String displayName = dimension.DisplayName;
                            String gvalue = "";
                            String gFormattedValue = "";
                            gFormattedValue = SolidEdgeUOM.FormatUnit(UOM, dimension, logFilePath);
                            gvalue = dimension.Value.ToString();

                            String Formula = dimension.Formula;
                            if (Formula != null && Formula.Equals("") == false)
                            {
                                if (Formula.StartsWith("@") == true)
                                {
                                    Utlity.Log("dimensionName: " + displayName, logFilePath);
                                    Utlity.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                    dimension.Formula = "";
                                }

                            }
                            //dimension.Formula = "";
                            if (gFormattedValue.Equals(updatedVar.value) == false)
                            {
                                //Utlity.Log("GetValue: " + gvalue, logFilePath);
                                String updatedValue = "";
                                //if (gvalue.Equals(dimension.Value) == false)
                                //{
                                //    updatedValue = SolidEdgeUOM.ParseUnit(UOM, updatedVar, logFilePath);
                                //}else {
                                updatedValue = updatedVar.value;
                                //}
                                //String updatedValue = updatedVar.value;
                                //Utlity.Log("updatedValue: " + updatedValue, logFilePath);
                                try
                                {
                                    if (updatedVar.UnitType.Equals("igDimTypeAngular", StringComparison.OrdinalIgnoreCase) == true ||
                                        updatedVar.UnitType.Equals("igDimTypeLinear", StringComparison.OrdinalIgnoreCase) == true ||
                                        updatedVar.UnitType.Equals("igDimTypeRDiameter", StringComparison.OrdinalIgnoreCase) == true ||
                                        updatedVar.UnitType.Equals("igDimTypeArcAngle", StringComparison.OrdinalIgnoreCase) == true ||
                                        updatedVar.UnitType.Equals("igDimTypeRadial", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        updatedValue = SolidEdgeUOM.ParseUnit(UOM, updatedVar, logFilePath);
                                    }
                                    else
                                    {
                                        updatedValue = StringUtils.RemoveUnitsFromDimension(updatedValue);
                                    }
                                    dimension.Value = double.Parse(updatedValue, System.Globalization.CultureInfo.InvariantCulture);
                                }
                                catch (Exception ex)
                                {
                                    Utlity.Log("Exception: " + ex.Message + "::::" + updatedValue, logFilePath);
                                }
                                //Utlity.Log("Setting Value: " + updatedValue, logFilePath);

                                Utlity.Log("DimName: " + dimension.DisplayName +
                                     "  DimFormattedValue: " + gFormattedValue +
                                 "  DimUserSpecifiedValue: " + updatedVar.value +
                                 "  DimTrimmedValue: " + updatedValue +
                                 "  DimCurrentSysValue: " + gvalue +
                                "  DimUpdatedSysValue: " + dimension.Value, logFilePath);
                            }

                            // 31/ AUG NOTE ---- DIMENSION UPDATE OF UPPER AND LOWER LIMIT IS NOT SUPPORTED IN SE API
                        }
                        break;
                    case SolidEdgeFramework.ObjectType.igVariable:
                        var variable = (SolidEdgeFramework.variable)item;
                        if (variable != null)
                        {
                            String displayName = variable.DisplayName;

                            String gvalue = "";
                            String gFormattedValue = "";
                            variable.GetValue(out gvalue);
                            gFormattedValue = SolidEdgeUOM.FormatUnit(UOM, variable, logFilePath);

                            String Formula = variable.Formula;
                            if (Formula != null && Formula.Equals("") == false)
                            {
                                if (Formula.StartsWith("@") == true)
                                {
                                    Utlity.Log("VariableName: " + displayName, logFilePath);
                                    Utlity.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                    variable.Formula = "";
                                }

                            }
                            //variable.Formula = "";
                            if (gFormattedValue.Equals(updatedVar.value) == false)
                            {
                                Utlity.Log("VariableName: " + displayName, logFilePath);
                                //Utlity.Log("GetValue: " + gvalue, logFilePath);
                                //String updatedValue = SolidEdgeUOM.ParseUnit(UOM, updatedVar, logFilePath);
                                String updatedValue = updatedVar.value;
                                try
                                {
                                    variable.SetValue(updatedValue);
                                }
                                catch (Exception ex)
                                {
                                    Utlity.Log("Exception: " + ex.Message + "::::" + updatedValue, logFilePath);
                                }

                                //Utlity.Log("Setting Value: " + updatedValue, logFilePath);
                                Utlity.Log("VarName: " + variable.DisplayName +
                                    "  VariableFormattedValue: " + gFormattedValue +
                                "  VarUserSpecifiedValue: " + updatedVar.value +
                                "  VarCurrentSysValue: " + gvalue +
                               "  VarUpdatedSysValue: " + variable.Value, logFilePath);
                            }

                            //UpdateBoundaries(variable, updatedVar,logFilePath);

                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("SetVariableData, Exception: " + ex.Message, logFilePath);
                return;
            }

        }

        private static void UpdateBoundaries(SolidEdgeFramework.variable variable, Variable updatedVar,String logFilePath)
        {
            if (variable != null && updatedVar != null)
            {
                try
                {
                    // ST- 8 Need to Check if this Works
                    variable.SetRange(updatedVar.rangeLow, updatedVar.rangeCondition, updatedVar.rangeHigh);
                    
                }
                catch (Exception ex)
                {
                    Utlity.Log("UpdateBoundaries: " + ex.Message, logFilePath);
                }

            }
        }

       

        private static void SaveAndCloseAssembly(SolidEdgeAssembly.AssemblyDocument assemblyDoc, String logFilePath)
        {
            try
            {
                if (assemblyDoc.ReadOnly == false)
                {
                    Utlity.Log("SaveAndCloseAssembly: " + " Save", logFilePath);
                    assemblyDoc.Save();
                    // 22 Sept
                    Utlity.Log("SaveAndCloseAssembly: " + "UpdateAll", logFilePath);
                    assemblyDoc.UpdateAll();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }
           

        }

        public static void traverseAssemblyToSetVariables(SolidEdgeAssembly.AssemblyDocument assemDoc, SolidEdgeFramework.Variables variables, String logFilePath, Dictionary<String, List<Variable>> variableDictionary, Dictionary<String, bool> partEnablementDictionary)
        {
            
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utlity.Log("variableDictionary is Empty ", logFilePath);
                return;
            }
            
            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                Utlity.Log("partEnablementDictionary is Empty ", logFilePath);
                return;
            }
            List<Variable> variablesList = null;
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            if (assemDoc == null)
            {
                Utlity.Log("assemDoc is Empty ", logFilePath);
                return;
            }
            variablesList = null;
            variableDictionary.TryGetValue(assemDoc.Name, out variablesList);
            if (variablesList == null || variablesList.Count == 0)
            {
                Utlity.Log("VariablesList is Empty--" + assemDoc.Name, logFilePath);

            }
            else if (variablesList.Count != 0)
            {
                

                foreach (Variable varr in variablesList)
                {
                    try
                    {
                        SetFormula(assemDoc.Name, varr.systemName, varr.Formula, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath, varr.variableType);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("SetFormula--" + ex.Message, logFilePath);
                    }
                    try
                    {
                        SetVariableData(assemDoc.UnitsOfMeasure, assemDoc.Name, varr.systemName, varr, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("SetVariableData--" + ex.Message, logFilePath);
                    }
                }
            }

            occurrences = assemDoc.Occurrences;
            if (occurrences == null)
            {
                Utlity.Log("occurrences is Empty " + assemDoc.Name, logFilePath);
                return;
            }
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
                


                SolidEdgePart.PartDocument partDoc = null;
                SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;
                SolidEdgeFramework.Variables variables1 = null;

                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                    variables1 = (SolidEdgeFramework.Variables)partDoc.Variables;
                    variablesList = null;
                    variableDictionary.TryGetValue(partDoc.Name, out variablesList);

                    if (variablesList == null || variablesList.Count == 0)
                    {
                        Utlity.Log("VariablesList is Empty--" + partDoc.Name, logFilePath);

                    }
                    else if (variablesList.Count != 0)
                    {
                        bool partEnable = false;
                        partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                        if (partEnable == false)
                        {
                            Utlity.Log("Part is NOT Enabled: " + occurenceName, logFilePath);
                            continue;
                        }

                        foreach (Variable varr in variablesList)
                        {
                           try {
                            SetFormula(partDoc.Name, varr.systemName, varr.Formula, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath,varr.variableType);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("SetFormula--" + ex.Message, logFilePath);
                            }
                           try
                           {
                               SetVariableData(partDoc.UnitsOfMeasure,partDoc.Name, varr.systemName, varr, variables1, logFilePath);
                           }
                           catch (Exception ex)
                           {
                               Utlity.Log("SetVariableData--" + ex.Message, logFilePath);
                           }
                        }
                    }

                    savePart(partDoc, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                    variables1 = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                    variablesList = null;
                    variableDictionary.TryGetValue(sheetMetalDoc.Name, out variablesList);
                    if (variablesList == null || variablesList.Count == 0)
                    {
                        Utlity.Log("VariablesList is Empty--" + sheetMetalDoc.Name, logFilePath);

                    }
                    else if (variablesList.Count != 0)
                    {
                        bool partEnable = false;
                        partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                        if (partEnable == false)
                        {
                            Utlity.Log("Part is NOT Enabled: " + occurenceName, logFilePath);
                            continue;
                        }

                        foreach (Variable varr in variablesList)
                        {
                            try {
                                SetFormula(sheetMetalDoc.Name, varr.systemName, varr.Formula, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, varr.variableType);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("SetFormula--" + ex.Message, logFilePath);
                            }

                            try
                            {
                                SetVariableData(sheetMetalDoc.UnitsOfMeasure,sheetMetalDoc.Name, varr.systemName, varr, variables1, logFilePath);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("SetVariableData--" + ex.Message, logFilePath);
                            }
                        }
                      }
                    saveSheet(sheetMetalDoc, logFilePath);
                    }                                    
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    variables1 = (SolidEdgeFramework.Variables)assemDoc.Variables;

                    traverseAssemblyToSetVariables(assemDoc1, variables1, logFilePath,variableDictionary,partEnablementDictionary);
                    SaveAndCloseAssembly(assemDoc1, logFilePath);
                }
        }


                //Utlity.Log("-----------------------------------------", logFilePath);
            }

        private static void savePart(SolidEdgePart.PartDocument partDoc, String logFilePath)
        {
            try
            {
                if (partDoc.ReadOnly == false)
                {
                    partDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }
        }

        private static void saveSheet(SolidEdgePart.SheetMetalDocument sheetMetalDoc, String logFilePath)
        {
            try
            {
                if (sheetMetalDoc.ReadOnly == false)
                {
                    sheetMetalDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }
        }
    }
}
