using ExcelSyncTC.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSyncTC.utils
{
    class SolidEdgeUOM
    {

        public static String StripUnitFromValue(String value, String logFilePath)
        {
            String Unit = "";

            String[] valueArray = value.Split(' ');
            if (valueArray != null && valueArray.Length == 2)
            {
                Unit = valueArray[1];
            }

            return Unit;
        }

        public static String StripValueWithoutUnits(String value, String logFilePath)
        {
            String Value = "";

            String[] valueArray = value.Split(' ');
            if (valueArray != null && valueArray.Length == 2)
            {
                Value = valueArray[0];
            }
            else
            {
                Value = value;
            }

            return Value;
        }

        public static String MergeValueAndUnit(String value, String Unit, String logFilePath)
        {
            String MergedValue = "";

            if (Unit != null && Unit.Equals("") == false)
            {
                MergedValue = value + " " + Unit;
            }
            else
            {
                MergedValue = value;
            }

            return MergedValue;
        }


        public static SolidEdgeFramework.UnitsOfMeasure getUOM(SolidEdgeAssembly.Occurrence occurrence, String logFilePath)
        {
            SolidEdgeFramework.UnitsOfMeasure uom = null;
            SolidEdgePart.PartDocument partDocument = null;
            SolidEdgePart.SheetMetalDocument sheetMetalDoc = null;
            SolidEdgeAssembly.AssemblyDocument assemDoc = null;
            try
            {
                if (occurrence.OccurrenceDocument is SolidEdgePart.PartDocument)
                {
                    partDocument = (SolidEdgePart.PartDocument)occurrence.OccurrenceDocument;
                    uom = partDocument.UnitsOfMeasure;
                }
                if (occurrence.OccurrenceDocument is SolidEdgePart.SheetMetalDocument)
                {
                    sheetMetalDoc = (SolidEdgePart.SheetMetalDocument)occurrence.OccurrenceDocument;
                    uom = sheetMetalDoc.UnitsOfMeasure;
                }

                if (occurrence.OccurrenceDocument is SolidEdgeAssembly.AssemblyDocument)
                {
                    assemDoc = (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument;
                    uom = assemDoc.UnitsOfMeasure;
                }   
            }
            catch (Exception ex)
            {
                Utlity.Log("getUOM: " + ex.Message, logFilePath);
            }
            return uom;
        }



        public static SolidEdgeFramework.UnitsOfMeasure getUOM(SolidEdgeFramework.SolidEdgeDocument document, String logFilePath)
        {
            SolidEdgeFramework.UnitsOfMeasure uom = null;
            
            if (document == null)
            {
                Utlity.Log("document is NULL: ", logFilePath);
                return null;
            }

            try
            {
                // Get a reference to the active document's unit of measure
                uom = document.UnitsOfMeasure;
            }
            catch (Exception ex)
            {
                Utlity.Log("getUOM: " + ex.Message,logFilePath);
            }
            return uom;
        }

        public static String FormatUnit(SolidEdgeFramework.UnitsOfMeasure uom, SolidEdgeFramework.variable variable, String logFilePath)
        {
            //SolidEdgeFramework.UnitsOfMeasure uom = getUOM(document,logFilePath);
            String sValue = "";
            if (uom == null)
            {
                Utlity.Log("uom is NULL: ", logFilePath);
                return "";
            }
            try
            {
                String dValue = (String)uom.FormatUnit((int)variable.UnitsType, variable.Value, Type.Missing);

                sValue = dValue.ToString();
            
            }
            catch (Exception ex)
            {
                Utlity.Log("FormatUnit: " + ex.Message, logFilePath);
            }

            return sValue;

        }


        public static String FormatUnit(SolidEdgeFramework.UnitsOfMeasure uom, SolidEdgeFrameworkSupport.Dimension dimension, String logFilePath)
        {
            
            String sValue = "";
            if (uom == null)
            {
                Utlity.Log("uom is NULL: ", logFilePath);
                return "";
            }
            try
            {
             
                String dimType = getType(dimension,logFilePath);

                if (dimType.Equals("igDimTypeRDiameter", StringComparison.OrdinalIgnoreCase) == true)
                {
                    String dValue = (String)uom.FormatUnit((int)SolidEdgeFramework.UnitTypeConstants.igUnitDistance, dimension.Value, Type.Missing); 
                    sValue = dValue.ToString();

                }
                else if (dimType.Equals("igDimTypeAngular", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //Double d = dimension.Value;
                    //double degree = RadianToDegree(d);
                    //sValue = degree.ToString();

                    String dValue = (String)uom.FormatUnit((int)SolidEdgeFramework.UnitTypeConstants.igUnitAngle, dimension.Value, Type.Missing);                  
                    sValue = dValue.ToString();
                }
                else if (dimType.Equals("igDimTypeLinear", StringComparison.OrdinalIgnoreCase) == true)
                {

                    String dValue = (String)uom.FormatUnit((int)SolidEdgeFramework.UnitTypeConstants.igUnitDistance, dimension.Value, Type.Missing);
                    
                    sValue = dValue.ToString();
                }
                else if (dimType.Equals("igDimTypeRadial", StringComparison.OrdinalIgnoreCase) == true)
                {

                    String dValue = (String)uom.FormatUnit((int)SolidEdgeFramework.UnitTypeConstants.igUnitDistance, dimension.Value, Type.Missing);
                    
                    sValue = dValue.ToString();
                }
                else if (dimType.Equals("igDimTypeArcAngle", StringComparison.OrdinalIgnoreCase) == true)
                {
                    // FIX on 20 - SEPT
                    //Double d = dimension.Value;
                    //double degree = RadianToDegree(d);
                    //sValue = degree.ToString();

                    String dValue = (String)uom.FormatUnit((int)SolidEdgeFramework.UnitTypeConstants.igUnitAngle, dimension.Value, Type.Missing);
                    sValue = dValue.ToString();
                }
                else
                {
                    sValue = dimension.Value.ToString();
                }


            }
            catch (Exception ex)
            {
                Utlity.Log("FormatUnit: " + ex.Message, logFilePath);
            }

            return sValue;

        }


        public static String ParseUnit(SolidEdgeFramework.UnitsOfMeasure uom, Variable varr, String logFilePath)
        {           
            String sValue = "";
            if (uom == null)
            {
                Utlity.Log("uom is NULL: ", logFilePath);
                return "";
            }
            try
            {
               
                if (varr.variableType.Equals("Var", StringComparison.OrdinalIgnoreCase) == true)
                {
                    int index = getIndexFromVariable(varr, logFilePath);
                    Double dValue = (Double)uom.ParseUnit(index, varr.value);
                    sValue = dValue.ToString();
                }
                else
                {
                    int index = getIndexFromVariable(varr, logFilePath);
                    sValue = ParseDimension(uom,varr);

                }

            }
            catch (Exception ex)
            {
                Utlity.Log("ParseUnit: " + ex.Message, logFilePath);
            }

            return sValue;

        }

        private static String ParseDimension(SolidEdgeFramework.UnitsOfMeasure uom, Variable varr)
        {
            String dimType = varr.UnitType;
            String sValue = "";
            if (dimType.Equals("igDimTypeRDiameter", StringComparison.OrdinalIgnoreCase) == true)
            {
                int index = (int)SolidEdgeFramework.UnitTypeConstants.igUnitDistance;
                Double dValue = (Double)uom.ParseUnit(index, varr.value);
                sValue = dValue.ToString();

            }
            else if (dimType.Equals("igDimTypeAngular", StringComparison.OrdinalIgnoreCase) == true)
            {
                int index = (int)SolidEdgeFramework.UnitTypeConstants.igUnitAngle;
                Double dValue = (Double)uom.ParseUnit(index, varr.value);
                sValue = dValue.ToString();
            }
            else if (dimType.Equals("igDimTypeLinear", StringComparison.OrdinalIgnoreCase) == true)
            {
                int index = (int)SolidEdgeFramework.UnitTypeConstants.igUnitDistance;
                Double dValue = (Double)uom.ParseUnit(index, varr.value);
                sValue = dValue.ToString();
                
            }
            else if (dimType.Equals("igDimTypeRadial", StringComparison.OrdinalIgnoreCase) == true)
            {
                int index = (int)SolidEdgeFramework.UnitTypeConstants.igUnitDistance;
                Double dValue = (Double)uom.ParseUnit(index, varr.value);
                sValue = dValue.ToString();
                
            }
            else if (dimType.Equals("igDimTypeArcAngle", StringComparison.OrdinalIgnoreCase) == true)
            {
                // 20 - SEPT - Included For Fix.
                int index = (int)SolidEdgeFramework.UnitTypeConstants.igUnitAngle;
                Double dValue = (Double)uom.ParseUnit(index, varr.value);
                sValue = dValue.ToString();

            }
            else
            {
                sValue = varr.value;
            }
            return sValue;
            
        }

        private static int getIndexFromVariable(Variable varr,String logFilePath)
        {
            int EnumIndex = 0;
            try
            {
                //Utlity.Log("ParseUnit: " + "varr.variableType " + varr.variableType, logFilePath);
                //Utlity.Log("ParseUnit: " + "varr.UnitType " + varr.UnitType, logFilePath);
                if (varr.variableType.Equals("Var", StringComparison.OrdinalIgnoreCase) == true)
                {
                    EnumIndex = (int)Enum.Parse(typeof(SolidEdgeFramework.UnitTypeConstants), varr.UnitType);
                }
                else if (varr.variableType.Equals("Dim", StringComparison.OrdinalIgnoreCase) == true)
                {
                    EnumIndex = (int)Enum.Parse(typeof(SolidEdgeFrameworkSupport.DimTypeConstants), varr.UnitType);
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("getIndexFromVariable: " + ex.Message, logFilePath);
            }

            return EnumIndex;
            
        }


        private static double DegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        private static double RadianToDegree(double angle)
        {
            return angle * (180.0 / Math.PI);
        }

        
        public static String getType(SolidEdgeFramework.variable variable, String logFilePath)
        {
            String type = "";

            type = Enum.GetName(typeof(SolidEdgeFramework.UnitTypeConstants), variable.UnitsType);
            

            return type;
        }


        public static String getType(SolidEdgeFrameworkSupport.Dimension dimension, String logFilePath)
        {
            String type = "";

            
            type = Enum.GetName(typeof(SolidEdgeFrameworkSupport.DimTypeConstants), dimension.DimensionType);
            return type;
        }
    }
}
