using DemoAddInTC.model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace DemoAddInTC.utils
{
    class ConvertDataTableToList
    {

        //private static String getValue(DataRow row, String propName, String logFilePath)
        //{
        //    String value = "";
        //    try
        //    {
        //        value = (String)row[propName];
        //    }
        //    catch (Exception ex)
        //    {
        //        //Utlity.Log(ex.Message + ":::" + propName, logFilePath);
        //        return "";
        //    }
        //    return value;
        //}

        public static List<Variable> ConvertDataTableToVariablesList(DataTable ds, String logFilePath)
        {
            //1 - Convert the Updated DGV to Table/List.            
            List<Variable> variablesList = new List<Variable>();
            Utlity.Log(ds.Rows.Count.ToString(), logFilePath);

            foreach (DataRow row in ds.Rows)
            {
                //if (row.IsNewRow == true) continue;
                //if (row.DataBoundItem is Subro.Controls.GroupRow)
                //{
                //    continue;
                //}
                //else if (row.DataBoundItem is DataGridViewRow)
                //{
                //    Utlity.Log("DataGridViewRow", logFilePath);
                //}

                Variable varr = new Variable();
                //Utlity.Log((String)row["PartName"], logFilePath);
                varr.PartName = Utlity.getValue(row, "PartName", logFilePath);
                varr.name = Utlity.getValue(row, "name", logFilePath);
                varr.systemName = Utlity.getValue(row, "systemName", logFilePath);
                varr.value = Utlity.getValue(row, "value", logFilePath);
                varr.DefaultValue = varr.value;
                varr.unit = Utlity.getValue(row, "unit", logFilePath);
                varr.rangeLow = Utlity.getValue(row, "rangeLow", logFilePath);
                varr.rangeHigh = Utlity.getValue(row, "rangeHigh", logFilePath);
                try
                {
                    int result = 0;
                    Int32.TryParse(row["rangeCondition"].ToString(), out result);
                    varr.rangeCondition = result;

                }
                catch (Exception ex)
                {
                    Utlity.Log(ex.Message, logFilePath);
                }
                varr.Formula = Utlity.getValue(row, "Formula", logFilePath);
                varr.LOV = Utlity.getValue(row, "LOV", logFilePath);
                varr.AddVarToTemplate = Utlity.getBoolValue(row, "AddVarToTemplate", logFilePath);
                //varr.AddPartToTemplate = getValue(row, "AddPartToTemplate", logFilePath);
                varr.AddPartToTemplate = "Y";
                varr.variableType = Utlity.getValue(row, "variableType", logFilePath);
                varr.UnitType = Utlity.getValue(row, "UnitType", logFilePath);
                variablesList.Add(varr);
                //Utlity.Log(varr.systemName, logFilePath);
            }

            return variablesList;
        }

        public static List<FeatureLine> ConvertDataTableToFeaturesList(DataTable ds1,String logFilePath)
        {
            
            //1 - Convert the Updated DGV to Table/List.            
            List<FeatureLine> featureLineList = new List<FeatureLine>();
            Utlity.Log(ds1.Rows.Count.ToString(), logFilePath);

            foreach (DataRow row in ds1.Rows)
            {
                FeatureLine fl = new FeatureLine();
                fl.SystemName = Utlity.getValue(row, "SystemName", logFilePath);
                Utlity.Log("fl.SystemName: " + fl.SystemName, logFilePath);
                fl.FeatureName = Utlity.getValue(row, "FeatureName", logFilePath);
                Utlity.Log("fl.FeatureName: " + fl.FeatureName, logFilePath);
                fl.Formula = Utlity.getValue(row, "Formula", logFilePath);
                Utlity.Log("fl.Formula: " + fl.Formula, logFilePath);
                fl.PartName = Utlity.getValue(row, "PartName", logFilePath);
                Utlity.Log("fl.PartName: " + fl.PartName, logFilePath);
                // 14 - OCT - LTC Issue FIX V1.0.04
                String SuppressionEnabled = Utlity.getValue(row, "SuppressionEnabled", logFilePath);
                if (SuppressionEnabled != null && SuppressionEnabled.Equals("") == false)
                {
                    if (SuppressionEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        fl.IsFeatureEnabled = "N";
                    }
                    else
                    {
                        fl.IsFeatureEnabled = "Y";
                    }
                }
                else
                {
                    fl.IsFeatureEnabled = Utlity.getValue(row, "IsFeatureEnabled", logFilePath);
                }

                Utlity.Log("Utlity.getValue: " + Utlity.getValue(row, "SuppressionEnabled", logFilePath), logFilePath);
                Utlity.Log("fl.SuppressionEnabled: " + fl.SuppressionEnabled, logFilePath);
                Utlity.Log("fl.IsFeatureEnabled: " + fl.IsFeatureEnabled, logFilePath);
                fl.EdgeBarName = Utlity.getValue(row, "EdgeBarName", logFilePath);
                featureLineList.Add(fl);
            }

            return featureLineList;
        }
    }
}
