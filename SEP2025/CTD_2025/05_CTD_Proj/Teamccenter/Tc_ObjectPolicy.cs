using AddToTc.CDAT_BulkUploader;
using CreoToTc.Utils;
using Log;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Common;

namespace Creo_TC_Live_Integration.TeamCenter
{
    class ObjectPolicy
    {

        public static bool setObjectPolicy_General()
        {
            try
            {
                String[] props = new String[19];

                if (Tc_Services.ss == null)
                {
                    log.write(logType.ERROR, "Session object is null in set object policy function");

                    return false;
                }
                props[props.Count() - 19] = "based_on";
                props[props.Count() - 18] = "Project_name";
                props[props.Count() - 17] = "project_list";
                props[props.Count() - 16] = "projects_list";
                props[props.Count() - 15] = "project_ids";
                props[props.Count() - 14] = "checked_out_user";

                props[props.Count() - 13] = "revision_list";

                props[props.Count() - 12] = "IMAN_Rendering";

                props[props.Count() - 11] = "uid";

                props[props.Count() - 10] = "cd4InstanceIdCp";

                props[props.Count() - 9] = "release_statuses";

                props[props.Count() - 8] = "release_status_list";

                props[props.Count() - 7] = "current_revision_id";

                props[props.Count() - 6] = "checked_out";

                props[props.Count() - 5] = "isObsolete";

                props[props.Count() - 4] = "revision_list";

                props[props.Count() - 3] = "item_id";

                props[props.Count() - 2] = "item_revision_id";

                props[props.Count() - 1] = "IMAN_specification";

                ObjectPropertyPolicy policy = new ObjectPropertyPolicy();

                policy.AddType(new PolicyType("ItemRevision", props));

                policy.AddType(new PolicyType("UnitOfMeasure", new string[] { "symbol", "unit" }));

                policy.AddType(new PolicyType("Item", new string[] { "uom_tag", "object_type", "object_name", "item_id", "bom_view_tags", "item_revision_id", "IMAN_reference", "IMAN_specification", "revision_list", "bom_view_tags" }));

                policy.AddType(new PolicyType("BOMLine", new String[] {"bl_rev_item_revision_id", "bl_item_item_id", "ps_children", "Bl_child_lines", "bl_all_notes", "bl_bomview_rev", "bl_bomview", "CD4InstanceId", "cd4InstanceId", "bl_line_name", "bl_sequence_no", "bl_plmxml_abs_xform", "bl_all_notes", "bl_quantity", "SE ObjectID", "ps_children", "bl_rev_ps_children", "bl_all_child_lines", "bl_child_lines" }));

                policy.AddType(new PolicyType("BOMWindow", new String[] { "bl_rev_item_revision_id", "is_packed_by_default" }));

                policy.AddType(new PolicyType("TC_Project", new String[] { "project_name" }));

                policy.AddType("RevisionRule", new String[] { "object_name", "Rule_date" });

                policy.AddType("RevisionRuleInfo", new String[] { "object_name", "Rule_date" });

                policy.AddType("ReleaseStatus", new String[] { "Name", "name", "Object_name", "name" });

                policy.AddType("release_status_list", new String[] { "Object_name", "name" });

                policy.AddType("User", new String[] { "Userid", "User_id", "user_name", });

                policy.AddType("ImanFile", new String[] { "original_file_name" });

                policy.AddType(new PolicyType("Dataset", new String[] { "Checked_out_user", "date_released", "release_status_list","object_name", "type_name", "dataset", "original_file_name", "object_string", "cd4UploadTime", "object_desc", "ref_list", "checked_out", "user_name", "checked_out_user", "dataset_type", "datasettype_name", "object_type", "owning_user" }));

                policy.AddType(new PolicyType("DatasetType", new String[] { "Checked_out_user", "dataset", "ref_list", "original_file_name", "cd4UploadTime", "Dataset_type", "datasettype_name", "type_name", "checked_out", "checked_out_user", "user_name" }));

                policy.AddType(new PolicyType("EPMTask", new String[] { "Object_name", "Object_string" }));
                policy.AddType(new PolicyType("EPMTask", new String[] { "object_name", "object_string" }));
                policy.AddType(new PolicyType("EPMTaskTemplate", new String[] { "object_name", "object_string" }));
                policy.AddType(new PolicyType("EPMTaskTemplate", new String[] { "Object_name", "Object_string" }));

                PolicyType datasetType = new PolicyType("Dataset");

                PolicyProperty property = new PolicyProperty("dataset_type");

                property.SetModifier(PolicyProperty.WITH_PROPERTIES, true);

                datasetType.AddProperty(property);

                policy.AddType(datasetType);

                Tc_Services.ss.SetObjectPropertyPolicy(policy);

                return true;
            }

            catch (Exception e)
            {
                log.writeException(e, "General set part object policy");
                return false;
            }
        }

        internal static void set_BOM_Obj_Policy(Dictionary<string, Dictionary<String, String>> objType_Vs_prop_RealName_Vs_DisplayName)
        {
            try
            {
               // List<string> endValues = CollectEndValues(parentNode_ErpProp_Vs_TcRealName, childNode_ErpProp_Vs_TcRealName);

                List<string> itemRevisionProPRealName = get_Props_RealName(objType_Vs_prop_RealName_Vs_DisplayName, "ItemRev");
                List<string> bl_ProPRealName = get_Props_RealName(objType_Vs_prop_RealName_Vs_DisplayName, "BomLine");
                List<string> itemPropRealName = get_Props_RealName(objType_Vs_prop_RealName_Vs_DisplayName, "Item");
                bl_ProPRealName.Add("bl_child_lines");
                bl_ProPRealName.Add("bl_revision");
                bl_ProPRealName.Add("bl_item");

                String[] props = new String[itemRevisionProPRealName.Count() +16];

                itemRevisionProPRealName.CopyTo(props, 0);

                if (Tc_Services.ss == null)
                {
                    return;
                }

                props[props.Count() - 16] = "project_list";
                props[props.Count() - 15] = "projects_list";
                props[props.Count() - 14] = "checked_out_user";

                props[props.Count() - 13] = "revision_list";

                props[props.Count() - 12] = "IMAN_Rendering";

                props[props.Count() - 11] = "uid";

                props[props.Count() - 10] = "cd4InstanceIdCp";

                props[props.Count() - 9] = "release_statuses";

                props[props.Count() - 8] = "release_status_list";

                props[props.Count() - 7] = "current_revision_id";

                props[props.Count() - 6] = "checked_out";

                props[props.Count() - 5] = "isObsolete";

                props[props.Count() - 4] = "revision_list";

                props[props.Count() - 3] = "item_id";

                props[props.Count() - 2] = "item_revision_id";

                props[props.Count() - 1] = "IMAN_specification";

                ObjectPropertyPolicy policy = new ObjectPropertyPolicy();

                policy.AddType(new PolicyType("ItemRevision", props));

                policy.AddType(new PolicyType("ItemRevision", new string[] { "CD4_2D_3D_link", "Checked_out_user", "project_ids" }));
                policy.AddType(new PolicyType("UnitOfMeasure", new string[] { "symbol", "unit" }));

                policy.AddType(new PolicyType("Item", itemPropRealName.ToArray()));

                //  policy.AddType(new PolicyType("BOMLine", new String[] { "object_type", "bl_rev_item_revision_id", "bl_item_item_id", "ps_children", "Bl_child_lines", "bl_all_notes", "bl_bomview_rev", "bl_bomview", "CD4InstanceId", "cd4InstanceId", "bl_line_name", "bl_sequence_no", "bl_plmxml_abs_xform", "bl_all_notes", "bl_quantity", "SE ObjectID", "ps_children", "bl_rev_ps_children", "bl_all_child_lines", "bl_child_lines" }));

                policy.AddType(new PolicyType("BOMLine", bl_ProPRealName.ToArray()));

               policy.AddType(new PolicyType("BOMWindow", new String[] { "object_type", "bl_rev_item_revision_id", "is_packed_by_default" }));

                policy.AddType(new PolicyType("WorkspaceObject", new String[] { "item_revision", "object_type", "object_name", "bl_item_item_id", "bl_rev_item_revision_id" }));

                policy.AddType(new PolicyType("EPMTask", new String[] { "Object_name", "Object_string" }));
                policy.AddType(new PolicyType("EPMTask", new String[] { "object_name", "object_string" }));
                policy.AddType(new PolicyType("EPMTaskTemplate", new String[] { "object_name", "object_string" }));
                policy.AddType(new PolicyType("EPMTaskTemplate", new String[] { "Object_name", "Object_string" }));
                policy.AddType(new PolicyType("TC_Project", new String[] { "project_name" }));

                policy.AddType("RevisionRule", new String[] { "object_name", "Rule_date" });

                policy.AddType("RevisionRuleInfo", new String[] { "object_name", "Rule_date" });

                policy.AddType("ReleaseStatus", new String[] { "Name", "name", "Object_name", "name", "object_name" });

                policy.AddType("release_status_list", new String[] { "Object_name", "name" });

                policy.AddType("User", new String[] { "Userid", "User_id", "user_name" });

                policy.AddType("ImanFile", new String[] { "original_file_name" });

                policy.AddType(new PolicyType("Dataset", new String[] { "Checked_out_user", "date_released", "owning_user", "release_status_list", "object_name", "type_name", "dataset", "original_file_name", "object_string", "cd4UploadTime", "object_desc", "ref_list", "checked_out", "checked_out_user", "user_name", "dataset_type", "datasettype_name", "object_type" }));

                policy.AddType(new PolicyType("DatasetType", new String[] { "Checked_out_user", "checked_out", "checked_out_user", "user_name", "dataset", "ref_list", "original_file_name", "cd4UploadTime", "Dataset_type", "datasettype_name", "type_name" }));

                Tc_Services.ss.SetObjectPropertyPolicy(policy);
            }
            catch (Exception ex)
            {
                log.writeException(ex, "General set bom object policy");

            }
        }

        private static List<string> get_Props_RealName(Dictionary<string, Dictionary<string, string>> objType_Vs_prop_RealName_Vs_DisplayName,string type)
        {
            List<string> returnValue = new List<string>();
            try
            {
                foreach (var objectTypeEntry in objType_Vs_prop_RealName_Vs_DisplayName)
                {
                    if(objectTypeEntry.Key.CompareTo(type)!=0)
                    {
                        continue;
                    }
                    foreach (var propRealNameDisplayNameEntry in objectTypeEntry.Value)
                    {
                        returnValue.Add(propRealNameDisplayNameEntry.Key);
                    }

                    Console.WriteLine();
                }

                return returnValue;
            }

            catch (Exception ex)
            {
                log.writeException(ex, "get_Props_RealName");
                return returnValue;

            }
        }

        internal static void set_Part_Obj_Policy(Dictionary<string, List<Dictionary<string, string>>> nodeName_Vs_ErpPropName_Vs_TcPropName)
        {
            try
            {
                List<string> customProp = new List<string>();

                foreach (var key in nodeName_Vs_ErpPropName_Vs_TcPropName.Keys)
                {
                    foreach (var innerDict in nodeName_Vs_ErpPropName_Vs_TcPropName[key])
                    {
                        foreach (var value in innerDict.Values)
                        {
                            customProp.Add(value);
                        }
                    }
                }

                set_Obj_Policy(customProp);
            }
            catch (Exception ex)
            {
                log.writeException(ex, "General set part object policy");
            }
        }
        internal static void set_Obj_Policy(List<string> _props)
        {
            try
            {
             

                if (_props.Count == 0)
                {
                    setObjectPolicy_General();
                }

                String[] props = new String[_props.Count + 16];

                _props.CopyTo(props, 0);

                if (Tc_Services.ss == null)
                {
                    return;
                }
                props[props.Count() - 16] = "project_list";
                props[props.Count() - 15] = "projects_list";
                props[props.Count() - 14] = "checked_out_user";

                props[props.Count() - 13] = "revision_list";

                props[props.Count() - 12] = "IMAN_Rendering";

                props[props.Count() - 11] = "uid";

                props[props.Count() - 10] = "cd4InstanceIdCp";

                props[props.Count() - 9] = "release_statuses";

                props[props.Count() - 8] = "release_status_list";

                props[props.Count() - 7] = "current_revision_id";

                props[props.Count() - 6] = "checked_out";

                props[props.Count() - 5] = "isObsolete";

                props[props.Count() - 4] = "revision_list";

                props[props.Count() - 3] = "item_id";

                props[props.Count() - 2] = "item_revision_id";

                props[props.Count() - 1] = "IMAN_specification";

                ObjectPropertyPolicy policy = new ObjectPropertyPolicy();

                policy.AddType(new PolicyType("ItemRevision", props));

                policy.AddType(new PolicyType("ItemRevision", new string[] { "CD4_2D_3D_link", "Checked_out_user", "project_ids" }));
                policy.AddType(new PolicyType("UnitOfMeasure", new string[] { "symbol", "unit" }));

                policy.AddType(new PolicyType("Item", new string[] {  "uom_tag", "object_type", "object_name", "item_id", "bom_view_tags", "item_revision_id", "IMAN_reference", "IMAN_specification", "revision_list", "bom_view_tags" }));

                policy.AddType(new PolicyType("BOMLine", new String[] { "object_type", "bl_rev_item_revision_id", "bl_item_item_id", "ps_children", "Bl_child_lines", "bl_all_notes", "bl_bomview_rev", "bl_bomview", "CD4InstanceId", "cd4InstanceId", "bl_line_name", "bl_sequence_no", "bl_plmxml_abs_xform", "bl_all_notes", "bl_quantity", "SE ObjectID", "ps_children", "bl_rev_ps_children", "bl_all_child_lines", "bl_child_lines" }));

                policy.AddType(new PolicyType("BOMWindow", new String[] { "object_type", "bl_rev_item_revision_id", "is_packed_by_default" }));

                policy.AddType(new PolicyType("WorkspaceObject", new String[] { "item_revision", "object_type", "object_name", "bl_item_item_id", "bl_rev_item_revision_id" }));

                policy.AddType(new PolicyType("EPMTask", new String[] { "Object_name", "Object_string" }));
                policy.AddType(new PolicyType("EPMTask", new String[] { "object_name", "object_string" }));
                policy.AddType(new PolicyType("EPMTaskTemplate", new String[] { "object_name", "object_string" }));
                policy.AddType(new PolicyType("EPMTaskTemplate", new String[] { "Object_name", "Object_string" }));
                policy.AddType(new PolicyType("TC_Project", new String[] { "project_name" }));

                policy.AddType("RevisionRule", new String[] { "object_name", "Rule_date" });

                policy.AddType("RevisionRuleInfo", new String[] { "object_name", "Rule_date" });

                policy.AddType("ReleaseStatus", new String[] { "Name", "name", "Object_name", "name", "object_name" });

                policy.AddType("release_status_list", new String[] { "Object_name", "name" });

                policy.AddType("User", new String[] { "Userid", "User_id", "user_name" });

                policy.AddType("ImanFile", new String[] { "original_file_name" });

                policy.AddType(new PolicyType("Dataset", new String[] {  "Checked_out_user", "date_released", "owning_user", "release_status_list", "object_name", "type_name", "dataset", "original_file_name", "object_string", "cd4UploadTime", "object_desc", "ref_list", "checked_out", "checked_out_user", "user_name", "dataset_type", "datasettype_name", "object_type" }));

                policy.AddType(new PolicyType("DatasetType", new String[] {  "Checked_out_user", "checked_out", "checked_out_user", "user_name", "dataset", "ref_list", "original_file_name", "cd4UploadTime", "Dataset_type", "datasettype_name", "type_name" }));

                Tc_Services.ss.SetObjectPropertyPolicy(policy);

            }
            catch (Exception e)
            {
                log.writeException(e, "setObjPolicy");
              
            }
        }

        static List<string> CollectEndValues(Dictionary<string, string> parentNode, Dictionary<string, Dictionary<string, string>> childNode)
        {
            List<string> endValues = new List<string>();
            try
            {
               

                // Collect values from the parent dictionary
                foreach (var value in parentNode.Values)
                {
                    endValues.Add(value);
                }

                // Collect values from the child dictionary
                foreach (var childDict in childNode.Values)
                {
                    foreach (var value in childDict.Values)
                    {
                        endValues.Add(value);
                    }
                }

                return endValues;
            }
            catch (Exception ex)
            {
                log.writeException(ex, "CollectEndValues");
                return endValues;

            }
        }
    }
}
