using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC.opInterfaces
{
    class SolidEdgeInterface
    {
        public static List<String> occurenceList = new List<string>();

        public static void SolidEdgeSync(String assemblyFileName, String logFilePath,String Option)
        {
            occurenceList.Clear();
            Dictionary<String,List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utlity.Log("DEBUG - variableDictionary is Empty : " + assemblyFileName, logFilePath);
                return;
            }
            Dictionary<String, bool> partEnablementDictionary = ExcelData.getPartEnablementDictionary();
            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                Utlity.Log("DEBUG - partEnablementDictionary is Empty : " + assemblyFileName, logFilePath);
                return;
            }
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;
            
            try
            {
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
                    List<Variable>variablesList = null;
                    variableDictionary.TryGetValue(objAssemblyDocument.Name, out variablesList);

                    if (variablesList == null || variablesList.Count ==0)
                    {
                        Utlity.Log("VariablesList is Empty--" + objAssemblyDocument.Name, logFilePath);

                    }
                    else if (variablesList.Count != 0)
                    {
                        foreach (Variable varr in variablesList)
                        {
                            if (Option.Equals("VALUE", StringComparison.OrdinalIgnoreCase) == true)
                            {

                                SetVariableData(objAssemblyDocument.UnitsOfMeasure, objAssemblyDocument.Name,  varr, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath, "VALUE"); 
                                //SetVariables(objAssemblyDocument.Name, varr.systemName, varr.value, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath);
                            }
                            else if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                SetVariableData(objAssemblyDocument.UnitsOfMeasure, objAssemblyDocument.Name,  varr, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath, "DEFAULTVALUE"); 

                                //SetVariables(objAssemblyDocument.Name, varr.systemName, varr.DefaultValue, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath);

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

                        Utlity.Log("-----------------------------------------", logFilePath);
                        Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);

                        //02-12-2024 - Murali - SyncTEDialog.listOfFileNamesInSession some times misses child components - Reason not known.
                        //02-12-2024 - Murali - So if the child component is missed in the SEEC download logic, it can be added here to the list.
                        String ocurenceFileNameWithoutPath = Path.GetFileName(ocurenceFileName);
                        if(SyncTEDialog.listOfFileNamesInSession.Contains(ocurenceFileNameWithoutPath) == false)
                        {
                            Utlity.Log("listOfFileNamesInSession Adding Missed ocurence--" + ocurenceFileNameWithoutPath, logFilePath);
                            SyncTEDialog.listOfFileNamesInSession.Add(ocurenceFileNameWithoutPath);
                        }

                        //02-12-2024 - Murali - SyncTEDialog.listOfFileNamesInSession some times misses child components - Reason not known.
                        //02-12-2024 - Murali - So if the child component is missed in the SEEC download logic, it can be added here to the list.

                        SolidEdgePart.PartDocument partDoc = null;
                        SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                        SolidEdgeAssembly.AssemblyDocument assemDoc = null;
                        SolidEdgeFramework.Variables variables = null;

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            
                            partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)partDoc.Variables;
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
                                    Utlity.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                                    continue;
                                }

                                foreach (Variable varr in variablesList)
                                {
                                    if (Option.Equals("VALUE", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables, logFilePath);
                                        SetVariableData(partDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath, "VALUE");
                                    }
                                    else if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        //SetVariables(occurenceName, varr.systemName, varr.DefaultValue, (SolidEdgeFramework.Variables)variables, logFilePath);
                                        SetVariableData(partDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath,"DEFAULTVALUE");
                                    }
                                }
                            }
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                           

                            sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
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
                                    Utlity.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                                    continue;
                                }

                                foreach (Variable varr in variablesList)
                                {
                                    if (Option.Equals("VALUE", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        SetVariableData(sheetMetalDoc.UnitsOfMeasure, occurenceName,  varr, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, "VALUE");
                                        //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables, logFilePath);
                                    }
                                    else if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        //SetVariables(occurenceName, varr.systemName, varr.DefaultValue, (SolidEdgeFramework.Variables)variables, logFilePath);
                                        SetVariableData(sheetMetalDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, "DEFAULTVALUE");
                                    }
                                }
                            }
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                        {
                            assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)assemDoc.Variables;
                            traverseAssemblyToSetVariables(assemDoc, variables, logFilePath,Option);
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

            SaveAndCloseAssembly(objAssemblyDocument,logFilePath);

            occurenceList.Clear();
            partEnablementDictionary.Clear();
            variableDictionary.Clear();
            if (objApp != null) objApp.DisplayAlerts = true;  

        }

        public static void traverseAssemblyToSetVariables(SolidEdgeAssembly.AssemblyDocument assemDoc, SolidEdgeFramework.Variables variables, String logFilePath,String Option)
        {
            Dictionary<String, List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utlity.Log("variableDictionary is Empty ", logFilePath);
                return;
            }
            Dictionary<String, bool> partEnablementDictionary = ExcelData.getPartEnablementDictionary();
            if (partEnablementDictionary==null || partEnablementDictionary.Count == 0)
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

            }else if (variablesList.Count != 0)
            {
                foreach (Variable varr in variablesList)
                {                    
                    if (Option.Equals("VALUE", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        SetVariableData(assemDoc.UnitsOfMeasure, assemDoc.Name, varr, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath,"VALUE"); 
                        //SetVariables(assemDoc.Name, varr.systemName, varr.value, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath);
                    }
                    else if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        SetVariableData(assemDoc.UnitsOfMeasure, assemDoc.Name, varr, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath,"DEFAULTVALUE"); 
                        //SetVariables(assemDoc.Name, varr.systemName, varr.DefaultValue, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath);
                    }
                }
            }

            occurrences = assemDoc.Occurrences;
            if (occurrences == null)
            {
                Utlity.Log("occurrences is Empty " + assemDoc.Name, logFilePath);
                return;
            }
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

                Utlity.Log("-----------------------------------------", logFilePath);
                Utlity.Log("occurenceName--" + occurenceName, logFilePath);
                int occurenceQty = occurrence.Quantity;
                Utlity.Log("occurenceQty--" + occurenceQty, logFilePath);
                String ocurenceFileName = occurrence.OccurrenceFileName;
                Utlity.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);
                //bool partEnable = false;
                //partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                //if (partEnable == false)
                //{
                //    Utlity.Log("Part is NOT Enabled: " + occurenceName, logFilePath);
                //    continue;
                //}

                //02-12-2024 - Murali - SyncTEDialog.listOfFileNamesInSession some times misses child components - Reason not known.
                //02-12-2024 - Murali - So if the child component is missed in the SEEC download logic, it can be added here to the list.
                String ocurenceFileNameWithoutPath = Path.GetFileName(ocurenceFileName);
                if (SyncTEDialog.listOfFileNamesInSession.Contains(ocurenceFileNameWithoutPath) == false)
                {
                    Utlity.Log("listOfFileNamesInSession Adding Missed ocurence--" + ocurenceFileNameWithoutPath, logFilePath);
                    SyncTEDialog.listOfFileNamesInSession.Add(ocurenceFileNameWithoutPath);
                }
                //02-12-2024 - Murali - SyncTEDialog.listOfFileNamesInSession some times misses child components - Reason not known.
                //02-12-2024 - Murali - So if the child component is missed in the SEEC download logic, it can be added here to the list.

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
                            if (Option.Equals("VALUE", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                SetVariableData(partDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath,"VALUE");
                                //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables1, logFilePath);
                            }
                            else if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                SetVariableData(partDoc.UnitsOfMeasure, occurenceName,  varr, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath, "DEFAULTVALUE");
                                //SetVariables(occurenceName, varr.systemName, varr.DefaultValue, (SolidEdgeFramework.Variables)variables1, logFilePath);
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
                            if (Option.Equals("VALUE", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                SetVariableData(sheetMetalDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, "VALUE");
                                //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables1, logFilePath);
                            }
                            else if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                SetVariableData(sheetMetalDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, "DEFAULTVALUE");
                                //SetVariables(occurenceName, varr.systemName, varr.DefaultValue, (SolidEdgeFramework.Variables)variables1, logFilePath);
                            }
                        }
                    }
                    saveSheet(sheetMetalDoc, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    variables1 = (SolidEdgeFramework.Variables)assemDoc.Variables;

                    traverseAssemblyToSetVariables(assemDoc1, variables1, logFilePath,Option);
                    SaveAndCloseAssembly(assemDoc, logFilePath);
                }


                Utlity.Log("-----------------------------------------", logFilePath);
            }

        }

        

        private static void SetVariables(string occurenceName, String variableName,String value,SolidEdgeFramework.Variables variables, String logFilePath)
        {
            if (variables == null)
            {
                Utlity.Log("variables is NULL ", logFilePath);
                return;
            }
            if (value == null || value.Equals(""))
            {
                Utlity.Log("Value is NULL For: " + variableName, logFilePath);
                return;

            }

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
                        Utlity.Log("dimensionName: " + displayName, logFilePath);
                        String gvalue = "";
                        gvalue = dimension.Value.ToString();
                        Utlity.Log("GetValue: " + gvalue, logFilePath);
                        String Formula = dimension.Formula;
                        if (Formula != null && Formula.Equals("") == false)
                        {
                            if (Formula.StartsWith("@") == true)
                            {
                                Utlity.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                dimension.Formula = "";
                            }

                        }
                        //dimension.Formula = "";
                        if (gvalue.Equals(value) == false)
                        {
                            dimension.Value = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                            Utlity.Log("Setting Value: " + value, logFilePath);
                        }
                    }   
                    break;
                case SolidEdgeFramework.ObjectType.igVariable:
                    var variable = (SolidEdgeFramework.variable)item;
                    if (variable != null)
                    {
                        String displayName = variable.DisplayName;
                        Utlity.Log("VariableName: " + displayName, logFilePath);
                        String gvalue = "";
                        variable.GetValue(out gvalue);
                        Utlity.Log("GetValue: " + gvalue, logFilePath);
                        String Formula = variable.Formula;
                        if (Formula != null && Formula.Equals("") == false)
                        {
                            if (Formula.StartsWith("@") == true)
                            {
                                Utlity.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                variable.Formula = "";                            
                             }

                        }
                        //variable.Formula = "";
                        if (gvalue.Equals(value) == false)
                        {
                            variable.SetValue(value);
                            Utlity.Log("Setting Value: " + value, logFilePath);
                        }
                    }
                    break;
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

            try
            {
                //if (assemblyDoc.ReadOnly == false)
                //{
                //    assemblyDoc.Close(true);
                //}
            }
            catch (Exception ex)
            {
                Utlity.Log(ex.Message, logFilePath);
            }

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

        // 11 - SEPT - Modified for M to MM and back from MM to M conversion
        // System (Solid Edge) Stores data Only in M for Length and so On for other Unit Types
        // Hence its Needed to Convert the UserUnit to SystemUnit and Back while Setting and Reading Back

        private static void SetVariableData(SolidEdgeFramework.UnitsOfMeasure UOM, string occurenceName, Variable updatedVar, SolidEdgeFramework.Variables variables, String logFilePath,String Option)
        {
            String variableName = updatedVar.systemName;
            //Utlity.Log("variableName: " + variableName, logFilePath);
            String value = "";

            // 01- OCT- Changes for Requirement from LTC - Split Unit & Value
            updatedVar.value = SolidEdgeUOM.MergeValueAndUnit(updatedVar.value, updatedVar.unit, logFilePath);
            if (Option.Equals("DEFAULTVALUE", StringComparison.OrdinalIgnoreCase) == true)
            {
                value = updatedVar.DefaultValue;
            }
            else
            {
                value = updatedVar.value;
            }
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

            // 08-10-2024 | Murali | Fix for Issue Raised by Allen. If the Variable is set to FALSE in the XL, then it should not by synced at ALL.
            if (updatedVar.AddVarToTemplate == false)
            {
                Utlity.Log("updatedVar AddVarToTemplate is FALSE, So Returning..", logFilePath);
                return;
            }

            SolidEdgeFramework.VariableList variableList = null;
            string pFindCriterium = variableName;

            object NamedBy = SolidEdgeConstants.VariableNameBy.seVariableNameByBoth;
            object VarType = SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth;
            object CaseInsensitive = false;

            // Execute query.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
            if (variableList == null || variableList.Count == 0)
            {

                pFindCriterium = updatedVar.ToString();
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
                            //String updatedValue = SolidEdgeUOM.ParseUnit(UOM, updatedVar, logFilePath);
                            String updatedValue = updatedVar.value;
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
                        
                        // 31/ AUG NOTE ---- DIMENSION UPDATE OF UPPER AND LOWER LIMIT IS NOT SUPPORTED IN SE API (ST8).
                        // 14 SEPT -- SUPPORTED IN ST10
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
                            //Utlity.Log("VariableName: " + displayName, logFilePath);
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

                        //UpdateBoundaries(variable, updatedVar, logFilePath);

                    }
                    break;
            }

        }

        private static void UpdateBoundaries(SolidEdgeFramework.variable variable, Variable updatedVar, String logFilePath)
        {
            if (variable != null && updatedVar != null)
            {
                try                
                {                  
                    // ST9 Supports Range -- SetValueRangeLowValue & SetValueRangeHighValue
                    String UrangeLow = StringUtils.RemoveUnitsFromDimension(updatedVar.rangeLow);
                    String UrangeHigh = StringUtils.RemoveUnitsFromDimension(updatedVar.rangeHigh);
                    //String SysRangeLow = "";
                    //String SysRangeHigh = "";
                    //int Condition = 0;                                        
                    variable.SetRange(UrangeLow, updatedVar.rangeCondition, UrangeHigh);
                }
                catch (Exception ex)
                {
                    Utlity.Log("UpdateBoundaries: " + ex.Message, logFilePath);
                }

            }
        }

       

    }
}
