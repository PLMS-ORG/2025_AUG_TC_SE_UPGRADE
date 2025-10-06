using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace DemoAddInTC.controller
{
    class SolidEdgeFeatureVariableMapper
    {
        private static void WriteOutExtraProperties(SolidEdgeFrameworkSupport.Dimension dimension, string logFilePath)
        {
            Utlity.Log("VariableTableName: " + ":" + dimension.VariableTableName, logFilePath);
            object ParentObj = dimension.Parent;

            // Get the managed type.
            var type = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetType(ParentObj);
            Utlity.Log("Dimension - ParentObj :" + ":" + type.ToString(), logFilePath);

            SolidEdgeFramework.AttributeSets attributeSets = (SolidEdgeFramework.AttributeSets)dimension.AttributeSets;
            if (attributeSets != null)
            {
                foreach (SolidEdgeFramework.Attribute attribute in attributeSets)
                {
                    Utlity.Log("AttributeName:" + ":" + attribute.Name, logFilePath);
                    Utlity.Log("AttributeValue:" + ":" + attribute.Value, logFilePath);
                }
            }


            //---
            int Count = 0;
            dimension.GetRelatedCount(out Count);
            for (int i = 0; i < Count; i++)
            {
                object GraphicObject = null;
                double x;
                double y;
                double z;
                bool keypoint;

                dimension.GetRelated(i, out GraphicObject, out x, out y, out z, out keypoint);
                // Get the managed type.
                WriteFeatureDetails(GraphicObject, logFilePath);

            }

        }

        private static void WriteFeatureDetails(object GraphicObject, string logFilePath)
        {
            Type typ = GraphicObject.GetType(); // This returns type "System.__ComObject

            var type1 = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetType(GraphicObject);

            Utlity.Log("DimensionType:" + ":" + type1.ToString(), logFilePath);
            Utlity.Log("GraphicObject Name:" + ":" + GraphicObject.ToString(), logFilePath);

            if (type1.ToString().Contains("Reference") == true)
            {
                return;
            }

            // Show the name of the feature
            try
            {
                string name = (string)typ.InvokeMember("EdgebarName", BindingFlags.GetProperty, null, GraphicObject, null);
                Utlity.Log("EdgebarName:" + ":" + name, logFilePath);
            }
            catch (Exception ex)
            {
                Utlity.Log("EdgebarName: Could Not Retrieve the Details: " + ex.Message, logFilePath);
            }


            try
            {
                object Obj = (object)typ.InvokeMember("Parent", BindingFlags.GetProperty, null, GraphicObject, null);
                if (Obj != null)
                {
                    // Nasty Recursive Call
                    WriteFeatureDetails(Obj, logFilePath);
                }
            }
            catch (Exception ex)
            {
                Utlity.Log("object: Could Not Retrieve the Details: " + ex.Message, logFilePath);
                return;
            }



        }

        //private static void WriteOutExtraProperties(SolidEdgeFramework.variable variable, string logFilePath)
        //{
        //    Utlity.Log("VariableTableName: " + ":" + variable.VariableTableName, logFilePath);

        //    SolidEdgeFramework.AttributeSets attributeSets = (SolidEdgeFramework.AttributeSets)variable.AttributeSets;

        //    if (attributeSets == null) return;
        //    foreach (SolidEdgeFramework.Attribute attribute in attributeSets)
        //    {
        //        Utlity.Log("AttributeName:" + ":" + attribute.Name, logFilePath);
        //        Utlity.Log("AttributeValue:" + ":" + attribute.Value, logFilePath);
        //    }

        //}
    }
}
