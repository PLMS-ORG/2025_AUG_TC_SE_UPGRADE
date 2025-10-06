using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC
{
    class ViewUtils
    {

        public static void EnableDisableAllControls(bool enable, Control container)
        {
            foreach (Control c in container.Controls)
            {
                if (c is ProgressBar)
                {
                    continue;
                }
                if (c is Panel || c is GroupBox)
                {
                    EnableDisableAllControls(enable, c);
                }
                else
                {
                    c.Enabled = enable;
                }
            }

        }


        // Merge UI changes to SE List.
        public static List<Variable> MergeUserChangesToVariables(List<Variable> variablesList2, List<Variable> variablesList1, String logFilePath)
        {
            if (variablesList2 == null || variablesList2.Count == 0)
            {
                return null;
            }

            if (variablesList1 == null || variablesList1.Count == 0)
            {
                return variablesList2;
            }
            // SE Data.
            foreach (Variable var in variablesList1)
            {
                //Utlity.Log("var.systemName: " + var.systemName, logFilePath);
                Variable UIVariable = variablesList2.Find(varr => varr.systemName.Equals(var.systemName));
                if (UIVariable == null)
                {
                    Utlity.Log(" UIVariable could not be FOUND: " + var.systemName, logFilePath);
                    continue;
                }

                var.AddVarToTemplate = UIVariable.AddVarToTemplate;
                //Utlity.Log("UIVariable.AddVarToTemplate: " + UIVariable.AddVarToTemplate, logFilePath);
                var.LOV = UIVariable.LOV;
                //Utlity.Log("UIVariable.LOV: " + UIVariable.LOV, logFilePath);

            }


            return variablesList1;

        }
    }
}
