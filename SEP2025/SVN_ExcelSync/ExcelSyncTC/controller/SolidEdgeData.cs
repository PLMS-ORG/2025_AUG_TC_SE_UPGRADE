
using ExcelSyncTC.model;
using ExcelSyncTC.utils;
using SolidEdge.RevisionManager.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSyncTC.controller
{
    class SolidEdgeData
    {
        public static List<String> occurenceList = new List<string>();
        public static List<Variable> ALLvariablesList = new List<Variable>();
        public static Dictionary<String, List<Variable>> variableDictionary = new Dictionary<string, List<Variable>>
            ();
        public static List<BOMLine> bomLineList = new List<BOMLine>();
        public static Dictionary<String, String> ocurrencePathDictionary = new Dictionary<string, string>();
        public static String topLineAssemblyFileName;

        public SolidEdgeData()
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
            

            objDocuments = objApp.Documents;


            try
            {
                objAssemblyDocument = (SolidEdgeAssembly.AssemblyDocument)objDocuments.Open(assemblyFileName);
                Utlity.Log("DEBUG - InputFile is Opened : " + assemblyFileName, logFilePath);

                if (objAssemblyDocument.ReadOnly == true)
                {
                    bool WriteAccess = false;
                    objAssemblyDocument.SeekWriteAccess(out WriteAccess);
                    if (WriteAccess == false)
                    {
                        Utlity.Log("Could not get WriteAccess to--" + assemblyFileName, logFilePath);
                        return;
                    }
                }

                if (objAssemblyDocument != null)
                {
                    // This is for Top Assembly Alone
                    ReadAndFillVariables(objAssemblyDocument.Name, (SolidEdgeFramework.Variables)objAssemblyDocument.Variables, logFilePath);
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

                        SolidEdgePart.PartDocument partDoc = null;
                        SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
                        SolidEdgeAssembly.AssemblyDocument assemDoc = null;
                        SolidEdgeFramework.Variables variables = null;

                        if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                        {
                            partDoc = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)partDoc.Variables;
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                        {
                            sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)sheetMetalDoc.Variables;
                        }
                        else if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                        {
                            assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                            variables = (SolidEdgeFramework.Variables)assemDoc.Variables;
                        }

                        

                        ReadAndFillVariables(occurenceName, variables, logFilePath);

                        
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
                Utlity.Log("Exception: " + ex.Message, logFilePath);
                Utlity.Log("Exception: " + ex.Source, logFilePath);
            }

        }

        private static void ReadAndFillVariables(string occurenceName, SolidEdgeFramework.Variables variables,String logFilePath)
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
            List<Variable> variableArr = new List<Variable>();
            for (int j = 1; j <= variableList.Count; j++)
            {
                var item = variableList.Item(j);

                // Determine the runtime type of the object.
                var itemType = item.GetType();
                var objectType = (SolidEdgeFramework.ObjectType)itemType.InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, item, null);

                switch (objectType)
                {
                    case SolidEdgeFramework.ObjectType.igDimension:
                        var dimension = (SolidEdgeFrameworkSupport.Dimension)item;
                        if (dimension != null)
                        {
                            String displayName1 = dimension.DisplayName;
                            Utlity.Log("dimensionName: " + displayName1, logFilePath);
                            if (dimension.Expose == 1)
                            {
                                Utlity.Log("dimensionValue: " + dimension.Value, logFilePath);
                            }
                            Utlity.Log("dimension.ExposeName: " + dimension.ExposeName, logFilePath);
                            Utlity.Log("dimension.Formula: " + dimension.Formula, logFilePath);
                            Utlity.Log("dimension.Comment: " + dimension.GetComment(), logFilePath);
                        }
                        break;
                    case SolidEdgeFramework.ObjectType.igVariable:
                        var variable = (SolidEdgeFramework.variable)item;
                        if (variable != null)
                        {
                            String displayName = variable.DisplayName;
                            Utlity.Log("VariableName: " + displayName, logFilePath);
                            String value = "";
                            variable.GetValue(out value);
                            Utlity.Log("Variablevalue: " + value, logFilePath);
                            String lowValue = "";
                            int Condition;
                            String highValue = "";
                            variable.GetRange(out lowValue, out Condition, out highValue);
                            Utlity.Log("lowValue: " + lowValue, logFilePath);
                            Utlity.Log("Condition: " + Condition, logFilePath);
                            Utlity.Log("highValue: " + highValue, logFilePath);
                            Utlity.Log("variable.SystemName: " + variable.SystemName, logFilePath);
                            Utlity.Log("variable.Formula: " + variable.Formula, logFilePath);
                            try
                            {
                                Variable varr = new Variable();
                                varr.name = displayName;
                                varr.value = value;
                                varr.rangeLow = lowValue;
                                varr.rangeCondition = Condition;
                                varr.rangeHigh = highValue;
                                varr.systemName = variable.SystemName;
                                varr.PartName = occurenceName;
                                varr.Formula = variable.Formula;

                                ALLvariablesList.Add(varr);
                                variableArr.Add(varr);
                            }
                            catch (Exception ex)
                            {
                                Utlity.Log("Exception : " + ex.Message, logFilePath);
                                continue;
                            }

                        }
                        break;
                }
            }

            variableDictionary.Add(occurenceName, variableArr);
            
        }

        public static void traverseAssembly(String assemblyFileName, String logFilePath)
        {
            bomLineList.Clear();
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
            bomLineList.Add(blTop);

            linkDocuments = (SolidEdge.RevisionManager.Interop.LinkedDocuments)document.LinkedDocuments[RevisionManager.LinkTypeConstants.seLinkTypeNormal];

            for (int i = 1; i <= linkDocuments.Count; i++)
            {
                linkDocument = (SolidEdge.RevisionManager.Interop.Document)linkDocuments.Item[i];

                Utlity.Log("FullName: " + linkDocument.FullName, logFilePath);
                Utlity.Log("AbsolutePath: " + linkDocument.AbsolutePath, logFilePath);
                Utlity.Log("DocNum: " + linkDocument.DocNum, logFilePath);
                Utlity.Log("Revision: " + linkDocument.Revision, logFilePath);
                Utlity.Log("Status: " + linkDocument.Status, logFilePath);
                BOMLine bl = new BOMLine();
                bl.FullName = linkDocument.FullName;
                bl.AbsolutePath = linkDocument.AbsolutePath;
                bl.DocNum = linkDocument.DocNum;
                bl.Revision = linkDocument.Revision;
                bl.Status = getStatus(linkDocument);               
                bl.level = "1";
                bomLineList.Add(bl);
            }

            

            SE_SESSION.killRevisionManager(logFilePath);
           
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

    }
}
