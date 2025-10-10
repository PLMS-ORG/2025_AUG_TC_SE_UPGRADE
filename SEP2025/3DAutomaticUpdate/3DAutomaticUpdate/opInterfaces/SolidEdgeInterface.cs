using _3DAutomaticUpdate.controller;
using _3DAutomaticUpdate.model;
using _3DAutomaticUpdate.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _3DAutomaticUpdate.opInterfaces
{
    class SolidEdgeInterface
    {
        public static List<String> occurenceList = new List<string>();

        public static void SolidEdgeSync(String assemblyFileName, String logFilePath)
        {
            occurenceList.Clear();
            Dictionary<String, List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utility.Log("DEBUG - variableDictionary is Empty : " + assemblyFileName, logFilePath);
                return;
            }
            Dictionary<String, bool> partEnablementDictionary = ExcelData.getPartEnablementDictionary();
            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                Utility.Log("DEBUG - partEnablementDictionary is Empty : " + assemblyFileName, logFilePath);
                return;
            }
            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
            {
                Utility.Log("DEBUG : Solid Edge Application Object is NULL - ", logFilePath);
                return;
            }

            SolidEdgeAssembly.AssemblyDocument objAssemblyDocument = null;

            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            try
            {
                objDocuments = objApp.Documents;
                objApp.DisplayAlerts = false;
                //objApp.Visible = false;
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objApp.ActiveDocument;                
                
                //objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                Utility.Log("DEBUG - InputFile is Opened : " + assemblyFileName, logFilePath);

                if (objAssemblyDocument.ReadOnly == true)
                {
                    bool WriteAccess = false;
                    FileInfo f = new FileInfo(assemblyFileName);
                    f.IsReadOnly = false;
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
                    List<Variable> variablesList = null;
                    variableDictionary.TryGetValue(objAssemblyDocument.Name, out variablesList);
                    if (variablesList == null || variablesList.Count == 0)
                    {
                        Utility.Log("VariablesList is Empty--" + objAssemblyDocument.Name, logFilePath);

                    }
                    else if (variablesList.Count != 0)
                    {
                        foreach (Variable varr in variablesList)
                        {
                            //SetVariables(objAssemblyDocument.Name,varr.systemName,varr.value, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath);
                            SetVariableData(objAssemblyDocument.UnitsOfMeasure, objAssemblyDocument.Name, varr, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath, "VALUE");

                        }
                    }
                    occurrences = objAssemblyDocument.Occurrences;
                    //Utility.Log("occurrences.Count: " + occurrences.Count, logFilePath);
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

                        Utility.Log("-----------------------------------------", logFilePath);
                        Utility.Log("occurenceName--" + occurenceName, logFilePath);
                        int occurenceQty = occurrence.Quantity;
                        //Utility.Log("occurenceQty--" + occurenceQty, logFilePath);
                        String ocurenceFileName = occurrence.OccurrenceFileName;
                        //Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);
                        //bool partEnable = false;
                        //partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                        //if (partEnable == false)
                        //{
                        //    Utility.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                        //    continue;
                        //}


                        SolidEdgePart.PartDocument partDoc = null;
                        SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                        SolidEdgeAssembly.AssemblyDocument assemDoc = null;
                        SolidEdgeFramework.Variables variables = null;

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                            FileInfo f = new FileInfo(partDoc.FullName);
                            f.IsReadOnly = false;


                            if (partDoc.ReadOnly == true)
                            {
                                bool WriteAccess = false;
                               
                                partDoc.SeekWriteAccess(out WriteAccess);
                                if (WriteAccess == false)
                                {
                                    //Utility.ResetAlerts(objApp, false, logFilePath);
                                    Utility.Log("Could not get WriteAccess to--" + partDoc.FullName, logFilePath);
                                    //continue;
                                    
                                }
                            }

                            variables = (SolidEdgeFramework.Variables)partDoc.Variables;
                            variablesList = null;
                            variableDictionary.TryGetValue(partDoc.Name, out variablesList);
                            if (variablesList == null || variablesList.Count == 0)
                            {
                                Utility.Log("VariablesList is Empty--" + partDoc.Name, logFilePath);

                            }
                            else if (variablesList.Count != 0)
                            {
                                bool partEnable = false;
                                partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                                if (partEnable == false)
                                {
                                    Utility.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                                    continue;
                                }

                                foreach (Variable varr in variablesList)
                                {
                                    //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables, logFilePath);
                                    SetVariableData(partDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath, "VALUE");
                                }

                                savePart(partDoc, logFilePath);
                            }
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                            FileInfo f = new FileInfo(sheetMetalDoc.FullName);
                            f.IsReadOnly = false;
                            if (sheetMetalDoc.ReadOnly == true)
                            {
                                bool WriteAccess = false;
                                sheetMetalDoc.SeekWriteAccess(out WriteAccess);
                                if (WriteAccess == false)
                                {
                                    //Utility.ResetAlerts(objApp, false, logFilePath);
                                    Utility.Log("Could not get WriteAccess to--" + sheetMetalDoc.FullName, logFilePath);
                                    //continue;
                                    

                                }
                            }
                            
                            variables = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                            variablesList = null;
                            variableDictionary.TryGetValue(sheetMetalDoc.Name, out variablesList);
                            if (variablesList == null || variablesList.Count == 0)
                            {
                                Utility.Log("VariablesList is Empty--" + sheetMetalDoc.Name, logFilePath);
                            }
                            else if (variablesList.Count != 0)
                            {
                                bool partEnable = false;
                                partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                                if (partEnable == false)
                                {
                                    Utility.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                                    continue;
                                }

                                foreach (Variable varr in variablesList)
                                {
                                    //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables, logFilePath);
                                    SetVariableData(sheetMetalDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, "VALUE");
                                }
                                saveSheet(sheetMetalDoc, logFilePath);
                            }
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                        {
                            assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            FileInfo f = new FileInfo(assemDoc.FullName);
                            f.IsReadOnly = false;

                            if (assemDoc.ReadOnly == true)
                            {
                                bool WriteAccess = false;
                                assemDoc.SeekWriteAccess(out WriteAccess);
                                if (WriteAccess == false)
                                {
                                    //Utility.ResetAlerts(objApp, false, logFilePath);
                                    Utility.Log("Could not get WriteAccess to--" + assemDoc.FullName, logFilePath);
                                    //continue;
                                   

                                }
                            }

                            variables = (SolidEdgeFramework.Variables)assemDoc.Variables;
                            variablesList = null;
                            variableDictionary.TryGetValue(assemDoc.Name, out variablesList);
                            traverseAssemblyToSetVariables(assemDoc, variables, logFilePath);
                        }


                        //Utility.Log("-----------------------------------------", logFilePath);
                    }
                }
                else
                {
                    Utility.ResetAlerts(objApp, false, logFilePath);
                    return;
                }

            }
            catch (Exception ex)
            {
                Utility.Log("Exception: " + ex.Message, logFilePath);
                Utility.Log("Exception: " + ex.Source, logFilePath);
                Utility.Log("Exception: " + ex.StackTrace, logFilePath);
            }

            SaveAndCloseAssembly(objAssemblyDocument, logFilePath);

            occurenceList.Clear();
            partEnablementDictionary.Clear();
            variableDictionary.Clear();
            if (objApp != null) objApp.DisplayAlerts = false;

        }

        public static void traverseAssemblyToSetVariables(SolidEdgeAssembly.AssemblyDocument assemDoc, SolidEdgeFramework.Variables variables, String logFilePath)
        {
            Dictionary<String, List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
            if (variableDictionary == null || variableDictionary.Count == 0)
            {
                Utility.Log("variableDictionary is Empty ", logFilePath);
                return;
            }
            Dictionary<String, bool> partEnablementDictionary = ExcelData.getPartEnablementDictionary();
            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                Utility.Log("partEnablementDictionary is Empty ", logFilePath);
                return;
            }
            List<Variable> variablesList = null;
            SolidEdgeAssembly.Occurrences occurrences = null;
            SolidEdgeAssembly.Occurrence occurrence = null;

            if (assemDoc == null)
            {
                Utility.Log("assemDoc is Empty ", logFilePath);
                return;
            }
            variablesList = null;
            variableDictionary.TryGetValue(assemDoc.Name, out variablesList);
            if (variablesList == null || variablesList.Count == 0)
            {
                Utility.Log("VariablesList is Empty--" + assemDoc.Name, logFilePath);

            }
            else if (variablesList.Count != 0)
            {
                foreach (Variable varr in variablesList)
                {
                    //SetVariables(assemDoc.Name, varr.systemName, varr.value, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath);
                    SetVariableData(assemDoc.UnitsOfMeasure, assemDoc.Name, varr, (SolidEdgeFramework.Variables)assemDoc.Variables, logFilePath, "VALUE");
                }
                SaveAndCloseAssembly(assemDoc, logFilePath);
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
                //Utility.Log("ocurenceFileName--" + ocurenceFileName, logFilePath);
                //bool partEnable = false;
                //partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                //if (partEnable == false)
                //{
                //    Utility.Log("Part is NOT Enabled: " + occurenceName, logFilePath);
                //    continue;
                //}


                SolidEdgePart.PartDocument partDoc = null;
                SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                SolidEdgeAssembly.AssemblyDocument assemDoc1 = null;
                SolidEdgeFramework.Variables variables1 = null;

                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                    FileInfo f = new FileInfo(partDoc.FullName);
                    f.IsReadOnly = false;

                    if (partDoc.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        partDoc.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {                            
                            Utility.Log("Could not get WriteAccess to--" + partDoc.FullName, logFilePath);
                            
                            //continue;
                        }
                    }

                    variables1 = (SolidEdgeFramework.Variables)partDoc.Variables;
                    variablesList = null;
                    variableDictionary.TryGetValue(partDoc.Name, out variablesList);

                    if (variablesList == null || variablesList.Count == 0)
                    {
                        Utility.Log("VariablesList is Empty--" + partDoc.Name, logFilePath);

                    }
                    else if (variablesList.Count != 0)
                    {
                        bool partEnable = false;
                        partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                        if (partEnable == false)
                        {
                            Utility.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                            continue;
                        }

                        foreach (Variable varr in variablesList)
                        {
                            SetVariableData(partDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)partDoc.Variables, logFilePath, "VALUE");
                            //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables1, logFilePath);
                        }
                    }
                    savePart(partDoc, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;

                    FileInfo f = new FileInfo(sheetMetalDoc.FullName);
                    f.IsReadOnly = false;

                    if (sheetMetalDoc.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        sheetMetalDoc.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + sheetMetalDoc.FullName, logFilePath);
                            
                            //continue;
                        }
                    }

                    variables1 = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                    variablesList = null;
                    variableDictionary.TryGetValue(sheetMetalDoc.Name, out variablesList);
                    if (variablesList == null || variablesList.Count == 0)
                    {
                        Utility.Log("VariablesList is Empty--" + sheetMetalDoc.Name, logFilePath);

                    }
                    else if (variablesList.Count != 0)
                    {
                        bool partEnable = false;
                        partEnablementDictionary.TryGetValue(occurenceName, out partEnable);
                        if (partEnable == false)
                        {
                            Utility.Log("Skipping Sync of occurenceName--" + occurenceName, logFilePath);
                            continue;
                        }
                        foreach (Variable varr in variablesList)
                        {
                            //SetVariables(occurenceName, varr.systemName, varr.value, (SolidEdgeFramework.Variables)variables1, logFilePath);
                            SetVariableData(sheetMetalDoc.UnitsOfMeasure, occurenceName, varr, (SolidEdgeFramework.Variables)sheetMetalDoc.Variables, logFilePath, "VALUE");
                        }
                    }
                    saveSheet(sheetMetalDoc, logFilePath);
                }
                else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    assemDoc1 = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    FileInfo f = new FileInfo(assemDoc1.FullName);
                    f.IsReadOnly = false;

                    if (assemDoc1.ReadOnly == true)
                    {
                        bool WriteAccess = false;
                        assemDoc1.SeekWriteAccess(out WriteAccess);
                        if (WriteAccess == false)
                        {
                            Utility.Log("Could not get WriteAccess to--" + assemDoc1.FullName, logFilePath);
                            //continue;
                            
                        }
                    }

                    variables1 = (SolidEdgeFramework.Variables)assemDoc1.Variables;

                    traverseAssemblyToSetVariables(assemDoc1, variables1, logFilePath);
                    SaveAndCloseAssembly(assemDoc1, logFilePath);
                }


                Utility.Log("-----------------------------------------", logFilePath);
            }

        }

        private static void SetVariables(string occurenceName, String variableName, String value, SolidEdgeFramework.Variables variables, String logFilePath)
        {
            if (variables == null)
            {
                Utility.Log("variables is NULL ", logFilePath);
                return;
            }
            if (value == null || value.Equals(""))
            {
                Utility.Log("Value is NULL For: " + variableName, logFilePath);
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
                Utility.Log("variableList is NULL For: " + variableName, logFilePath);
                return;
            }
            var item = variableList.Item(1);

            // Determine the runtime type of the object.
            var itemType = item.GetType();
            if (itemType == null)
            {
                Utility.Log("itemType is NULL For: " + variableName, logFilePath);
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
                        Utility.Log("dimensionName: " + displayName, logFilePath);
                        String gvalue = "";
                        gvalue = dimension.Value.ToString();
                        Utility.Log("GetValue: " + gvalue, logFilePath);
                        String Formula = dimension.Formula;
                        if (Formula != null && Formula.Equals("") == false)
                        {
                            if (Formula.StartsWith("@") == true)
                            {
                                Utility.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                dimension.Formula = "";
                            }

                        }
                        //dimension.Formula = "";
                        if (gvalue.Equals(value) == false)
                        {
                            dimension.Value = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                            Utility.Log("Setting Value: " + value, logFilePath);
                        }
                    }
                    break;
                case SolidEdgeFramework.ObjectType.igVariable:
                    var variable = (SolidEdgeFramework.variable)item;
                    if (variable != null)
                    {
                        String displayName = variable.DisplayName;
                        Utility.Log("VariableName: " + displayName, logFilePath);
                        String gvalue = "";
                        variable.GetValue(out gvalue);
                        Utility.Log("GetValue: " + gvalue, logFilePath);
                        // Fix to set Formula to Empty Before setting the Value.
                        String Formula = variable.Formula;
                        if (Formula != null && Formula.Equals("") == false)
                        {
                            if (Formula.StartsWith("@") == true)
                            {
                                Utility.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                variable.Formula = "";
                            }

                        }
                        variable.Formula = "";

                        if (gvalue.Equals(value) == false)
                        {
                            variable.SetValue(value);
                            Utility.Log("Setting Value: " + value, logFilePath);
                        }
                    }
                    break;
            }




        }


        private static SolidEdgeFramework.Documents OpenDocument(String logFilePath)
        {


            SolidEdgeFramework.Documents objDocuments = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }



            objDocuments = objApp.Documents;

            if (objDocuments == null)
            {
                Utility.Log("objDocuments object is NULL", logFilePath);
                return null;

            }
            return objDocuments;
        }

        private static SolidEdgePart.PartDocument getPartDocument(SolidEdgeFramework.Documents objDocuments, String occurenceFilePath, String logFilePath)
        {
            SolidEdgePart.PartDocument partDoc = null;

            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }
            //objApp.DisplayAlerts = false;
            try
            {
                partDoc = (SolidEdgePart.PartDocument)objDocuments.Open(occurenceFilePath);
            }
            catch (Exception ex)
            {
                return null;
            }

            return partDoc;

        }

        private static SolidEdgeAssembly.AssemblyDocument getAssemblyDocument(SolidEdgeFramework.Documents objDocuments, String occurenceFilePath, String logFilePath)
        {
            SolidEdgeAssembly.AssemblyDocument assemDoc = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }

            try
            {
                //objApp.DisplayAlerts = false;
                assemDoc = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(occurenceFilePath);
            }
            catch (Exception ex)
            {
                return null;
            }
            return assemDoc;

        }

        private static SolidEdgePart.SheetMetalDocument getSheetMetalDocument(SolidEdgeFramework.Documents objDocuments, String occurenceFilePath, String logFilePath)
        {
            SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;

            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();
            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }

            try
            {
                //objApp.DisplayAlerts = false;
                sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)objDocuments.Open(occurenceFilePath);
            }
            catch (Exception ex)
            {
                return null;
            }


            return sheetMetalDoc;


        }

        private static SolidEdgeFramework.Variables GetVariablesForAssembly(SolidEdgeAssembly.AssemblyDocument assemDoc, String logFilePath)
        {


            SolidEdgeFramework.Variables variables = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }

            {
                //objApp.DisplayAlerts = false;                
                variables = (SolidEdgeFramework.Variables)assemDoc.Variables;
            }
            return variables;

        }

        private static SolidEdgeFramework.Variables GetVariablesForPart(SolidEdgePart.PartDocument partDoc, String logFilePath)
        {
            //SolidEdgePart.PartDocument partDoc = null;

            SolidEdgeFramework.Variables variables = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }


            {
                //objApp.DisplayAlerts = false;                
                variables = (SolidEdgeFramework.Variables)partDoc.Variables;

            }

            return variables;

        }

        private static SolidEdgeFramework.Variables GetVariablesForSM(SolidEdgePart.SheetMetalDocument sheetMetalDoc, String logFilePath)
        {
            SolidEdgeFramework.Variables variables = null;
            SolidEdgeFramework.Application objApp = SE_SESSION.getSolidEdgeSession();

            if (objApp == null)
            {
                Utility.Log("Application object is NULL", logFilePath);
                return null;
            }

            {
                //objApp.DisplayAlerts = false;                
                variables = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
            }

            return variables;

        }

        public static void setVariableInSolidEdge(SolidEdgeFramework.Variables variables, String variableName, String Value, String logFilePath)
        {
            SolidEdgeFramework.VariableList variableList = null;

            string pFindCriterium = "*" + variableName + "*";
            object NamedBy = SolidEdgeConstants.VariableNameBy.seVariableNameByBoth;
            object VarType = SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth;
            object CaseInsensitive = false;

            if (variables == null)
            {
                Utility.Log("variables object is NULL", logFilePath);
                return;
            }

            // Execute query.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);

            Utility.Log("variableList.Count:" + variableList.Count, logFilePath);

            for (int j = 1; j <= variableList.Count; j++)
            {
                var item = variableList.Item(j);

                // Determine the runtime type of the object.
                var itemType = item.GetType();
                var objectType = (SolidEdgeFramework.ObjectType)itemType.InvokeMember("Type",
                    System.Reflection.BindingFlags.GetProperty, null, item, null);

                switch (objectType)
                {
                    case SolidEdgeFramework.ObjectType.igDimension:
                        {
                            var dimension = (SolidEdgeFrameworkSupport.Dimension)item;
                            break;
                        }
                    case SolidEdgeFramework.ObjectType.igVariable:
                        {
                            var variable = (SolidEdgeFramework.variable)item;
                            if (variable != null)
                            {
                                variable.SetValue(Value);

                            }
                            break;
                        }
                }
            }

            // Save 



        }

        private static void SaveAndCloseSM(SolidEdgePart.SheetMetalDocument sheetMetalDoc, String logFilePath)
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
                Utility.Log(ex.Message, logFilePath);
            }

            try
            {
                if (sheetMetalDoc.ReadOnly == false)
                {
                    sheetMetalDoc.Close(true);
                }
            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
            }

        }

        private static void SaveAndClosePart(SolidEdgePart.PartDocument partDoc, String logFilePath)
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
                Utility.Log(ex.Message, logFilePath);
            }

            try
            {
                if (partDoc.ReadOnly == false)
                {
                    partDoc.Close(true);
                }
            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
            }

        }

        private static void SaveAndCloseAssembly(SolidEdgeAssembly.AssemblyDocument assemblyDoc, String logFilePath)
        {
            try
            {
                if (assemblyDoc.ReadOnly == false)
                {
                    Utility.Log("SaveAndCloseAssembly: " + " Save: " + assemblyDoc.FullName, logFilePath);
                    assemblyDoc.Save();
                    // 22 Sept
                    Utility.Log("SaveAndCloseAssembly: " + "UpdateAll: " + assemblyDoc.FullName, logFilePath);
                    assemblyDoc.UpdateAll();
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


        private static void savePart(SolidEdgePart.PartDocument partDoc, String logFilePath)
        {
            try
            {
                Utility.Log("Saving Editable Part: " + partDoc.FullName, logFilePath);
                //if (partDoc.ReadOnly == false)
                {
                    Utility.Log("Saving Part: " + partDoc.FullName, logFilePath);
                    partDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
            }
        }

        private static void saveSheet(SolidEdgePart.SheetMetalDocument sheetMetalDoc, String logFilePath)
        {
            try
            {
                //if (sheetMetalDoc.ReadOnly == false)
                {
                    Utility.Log("Save Sheet: " + sheetMetalDoc.FullName, logFilePath);
                    sheetMetalDoc.Save();
                }
            }
            catch (Exception ex)
            {
                Utility.Log(ex.Message, logFilePath);
            }
        }

        // 11 - SEPT - Modified for M to MM and back from MM to M conversion
        // System (Solid Edge) Stores data Only in M for Length and so On for other Unit Types
        // Hence its Needed to Convert the UserUnit to SystemUnit and Back while Setting and Reading Back

        private static void SetVariableData(SolidEdgeFramework.UnitsOfMeasure UOM, string occurenceName, Variable updatedVar, SolidEdgeFramework.Variables variables, String logFilePath, String Option)
        {
            String variableName = updatedVar.systemName;
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
                Utility.Log("variables is NULL ", logFilePath);
                return;
            }
            if (updatedVar == null)
            {
                Utility.Log("updatedVar is NULL ", logFilePath);
                return;

            }

            // 08-10-2024 | Murali | Fix for Issue Raised by Allen. If the Variable is set to FALSE in the XL, then it should not by synced at ALL.
            if (updatedVar.AddVarToTemplate == false)
            {
                Utility.Log("updatedVar AddVarToTemplate is FALSE, So Returning..", logFilePath);
                return;
            }

            SolidEdgeFramework.VariableList variableList = null;
            string pFindCriterium = variableName; // removed asterik based search - 05/04/2019 - Issue Raised by Simone

            object NamedBy = SolidEdgeConstants.VariableNameBy.seVariableNameByBoth;
            object VarType = SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth;
            object CaseInsensitive = false;

            // Execute query.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
            if (variableList == null || variableList.Count == 0)
            {

                pFindCriterium = updatedVar.ToString(); ; // removed asterik based search - 05/04/2019 - Issue Raised by Simone
                variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
                if (variableList == null || variableList.Count == 0)
                {
                    Utility.Log(updatedVar.name + " not Found..", logFilePath);
                    //pFindCriterium = "*";
                    //variableList = (SolidEdgeFramework.VariableList)variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive);
                    //if (variableList == null || variableList.Count == 0)
                    //{
                    //    Utility.Log("variableList is NULL For: " + updatedVar.name, logFilePath);
                    //    return;
                    //}
                    //else
                    //{
                    //    Utility.Log("variableList COUNT: " + variableList.Count, logFilePath);
                    //    return;
                    //}
                    return;
                }
                else
                {
                    Utility.Log("variableList COUNT: " + variableList.Count, logFilePath);
                }

            }


            try
            {
                var item = variableList.Item(1);
                if (item == null) return;
                // Determine the runtime type of the object.
                var itemType = item.GetType();
                if (itemType == null)
                {
                    Utility.Log("itemType is NULL For: " + variableName, logFilePath);
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
                                    Utility.Log("dimensionName: " + displayName, logFilePath);
                                    Utility.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                    dimension.Formula = "";
                                }

                            }
                            //dimension.Formula = "";

                            if (gFormattedValue.Equals(updatedVar.value) == false)
                            {
                                //Utility.Log("dimensionName: " + updatedVar.name, logFilePath);
                                //Utility.Log("GetValue: " + gvalue, logFilePath);
                                //String updatedValue = SolidEdgeUOM.ParseUnit(UOM, updatedVar, logFilePath);
                                //Utility.Log("updatedValue: " + updatedValue, logFilePath);
                                String updatedValue = updatedVar.value;
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

                                    Utility.Log("Exception: " + ex.Message + "::::" + updatedValue, logFilePath);
                                }

                                Utility.Log("DimName: " + dimension.DisplayName +
                                    "  DimFormattedValue: " + gFormattedValue +
                                "  DimUserSpecifiedValue: " + updatedVar.value +
                                "  DimCurrentSysValue: " + gvalue +
                               "  DimUpdatedSysValue: " + dimension.Value, logFilePath);
                                Utility.Log("Setting Value: " + updatedValue, logFilePath);
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
                                    Utility.Log("VariableName: " + displayName, logFilePath);
                                    Utility.Log("Setting Formula to Empty, Incase its a Excel Link: ", logFilePath);
                                    variable.Formula = "";
                                }

                            }
                            //variable.Formula = "";
                            if (gFormattedValue.Equals(updatedVar.value) == false)
                            {
                                //Utility.Log("VariableName: " + displayName, logFilePath);
                                //Utility.Log("GetValue: " + gvalue, logFilePath);
                                //String updatedValue = SolidEdgeUOM.ParseUnit(UOM, updatedVar, logFilePath);
                                String updatedValue = updatedVar.value;
                                variable.SetValue(updatedValue);
                                Utility.Log("VarName: " + variable.DisplayName +
                                   "  VariableFormattedValue: " + gFormattedValue +
                               "  VarUserSpecifiedValue: " + updatedVar.value +
                               "  VarCurrentSysValue: " + gvalue +
                              "  VarUpdatedSysValue: " + variable.Value, logFilePath);

                                Utility.Log("Setting Value: " + updatedValue, logFilePath);
                            }

                            //UpdateBoundaries(variable, updatedVar, logFilePath);

                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Utility.Log("SetVariableData, Exception: " + ex.Message, logFilePath);
                return;

            }

        }

        private static void UpdateBoundaries(SolidEdgeFramework.variable variable, Variable updatedVar, String logFilePath)
        {
            if (variable != null && updatedVar != null)
            {
                try
                {
                    variable.SetRange(updatedVar.rangeLow, updatedVar.rangeCondition, updatedVar.rangeHigh);
                }
                catch (Exception ex)
                {
                    Utility.Log("UpdateBoundaries: " + ex.Message, logFilePath);
                }

            }
        }

    }

}
