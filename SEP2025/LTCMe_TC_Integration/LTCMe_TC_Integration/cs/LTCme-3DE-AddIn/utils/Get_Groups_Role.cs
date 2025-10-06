using DemoAddInTC.se;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using Teamcenter.Services.Strong.Core._2006_03.Session;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;

//*** Set object policy

//policy.AddType("GroupMember", new String[] { "group", "the_group", "list_of_role" });

//policy.AddType("Group", new String[] {"display_name", "Display_name", "list_of_role", "List_of_role" });

//policy.AddType("Role", new String[] { "display_name", "role_name" });

//***** First call

namespace get_group_role_member
{

    class GroupRoleUserUtil
    {
        public static void get_User_Groups_And_Roles(string userID, string logfilepath, ref List<String> groupNames, ref List<string> roleNames)
        {
         //   groupNames = new List<string>();
          //  roleNames = new List<string>();

            try
            {             
                //**** Get Groups******

                List<Group> groups = get_Groups(userID,logfilepath);

                if (groups.Count == 0)
                {
                    Utility.Log("There is no group for : " + userID, logfilepath);
                    return;
                }

                //*******Get Rols******
               

                foreach (Group grp in groups)
                {
                    Utility.Log("Group : " + grp.Display_name, logfilepath);
                    if(!groupNames.Contains(grp.Display_name))
                         groupNames.Add(grp.Display_name);
                     //ServiceData sData = Tc_Services.dmService.LoadObjects(new string[] { grp.Uid });

                     ServiceData sData = TcAdaptor.dmService.LoadObjects(new string[] { grp.Uid });

                    for (int k = 0; k < sData.sizeOfPlainObjects(); k++)
                    {
                        Group _grp = (Group)sData.GetPlainObject(k);

                        Role[] _role = _grp.List_of_role;

                        foreach (Role rol in _role)
                        {
                            ServiceData rsData = TcAdaptor.dmService.LoadObjects(new string[] { rol.Uid });

                            for (int j = 0; j < rsData.sizeOfPlainObjects(); j++)
                            {
                                Role _rol = (Role)rsData.GetPlainObject(k);
                                Utility.Log("Role Name : " + _rol.Role_name, logfilepath);
                                if (roleNames.Contains(rol.Role_name))
                                    continue;
                                roleNames.Add(rol.Role_name);
                            }
                        }
                    }

                }

            }
            catch (System.Exception ex)
            {
                Utility.Log("exception in get_User_Groups_And_Roles :" + ex, logfilepath);
            }
        }

        //==================================================================================================================
        public static void get_User_Groups_And_Roles(string userID, string logfilepath, ref Dictionary<string,List<string>>groupVsRoleNames)
        {
            //   groupNames = new List<string>();
            //  roleNames = new List<string>();

            try
            {
                GetGroupMembershipResponse response = TcAdaptor.session.GetGroupMembership();


               GroupMember[] gmembers = response.GroupMembers;

                Utlity.Log("*************** Group and ROle ****************", logfilepath);

                foreach (GroupMember gm in gmembers)
                {
                    POM_group pomGroup = gm.Group;

                    string groupDisplayName = pomGroup.Display_name;

                    String _roleName=gm.Role.Role_name;

                    Utlity.Log("INFO : Group / Role Name : " + groupDisplayName+" / "+_roleName, logfilepath);

                    if (groupVsRoleNames.ContainsKey(groupDisplayName))
                    {
                        if (groupVsRoleNames[groupDisplayName].Contains(_roleName) == false)
                        {
                            groupVsRoleNames[groupDisplayName].Add(_roleName);
                        }
                    }
                    else
                    {
                        groupVsRoleNames[groupDisplayName]= new List<string>{ _roleName };
                    }
                }


                //**** Get Groups******

                //List<Group> groups = get_Groups(userID, logfilepath);

                //if (groups.Count == 0)
                //{
                //    Utlity.Log("There is no group for : " + userID, logfilepath);
                //    return;
                //}

                //*******Get Rols******

              /*  foreach (Group grp in groups)
                {
                    Utlity.Log("Group : " + grp.Display_name, logfilepath);

                    if(groupVsRoleNames.ContainsKey(grp.Display_name))
                        continue;

                    List<string> roleNames = new List<string>();
                    List<string> tempRoleNames;
                    if(groupVsRoleNames.TryGetValue(grp.Display_name, out tempRoleNames))
                        roleNames = tempRoleNames;

                    ServiceData sData = TcAdaptor.dmService.LoadObjects(new string[] { grp.Uid });

                    for (int k = 0; k < sData.sizeOfPlainObjects(); k++)
                    {
                        Group _grp = (Group)sData.GetPlainObject(k);

                        Role[] _role = _grp.List_of_role;

                        foreach (Role rol in _role)
                        {
                            ServiceData rsData = TcAdaptor.dmService.LoadObjects(new string[] { rol.Uid });

                            for (int j = 0; j < rsData.sizeOfPlainObjects(); j++)
                            {
                                Role _rol = (Role)rsData.GetPlainObject(k);
                                Utlity.Log("Role Name : " + _rol.Role_name, logfilepath);
                                if (roleNames.Contains(rol.Role_name))
                                    continue;
                                roleNames.Add(rol.Role_name);
                            }
                        }
                    }

                    groupVsRoleNames.Add(grp.Display_name, roleNames);
                }*/

            }
            catch (System.Exception ex)
            {
                Utlity.Log("exception in get_User_Groups_And_Roles :" + ex, logfilepath);
            }
        }

        //===============================================================================================
        //**** Get group data

        public static List<Group> get_Groups(string userID,string logfilepath)
        {

            List<Group> groups = new List<Group>();
            try
            {

                //if (Tc_Services.savedQryServices == null)
                //{
                //    log.write(logType.ERROR, "Save query service is not initialized");
                //    return groups;
                //}

               ImanQuery qry_EINT_group_members = Tc_Process.getTcQuery("__EINT_group_members",logfilepath);
              //  ImanQuery qry_EINT_group_members = Tc_Process.getTcQuery("__WEB_group", logfilepath);
                

                if (qry_EINT_group_members != null)
                {
                    Teamcenter.Services.Strong.Query._2008_06.SavedQuery.QueryInput[] qryInput = new Teamcenter.Services.Strong.Query._2008_06.SavedQuery.QueryInput[1];


                    qryInput[0] = new Teamcenter.Services.Strong.Query._2008_06.SavedQuery.QueryInput();

                    qryInput[0].Query = qry_EINT_group_members;

                    qryInput[0].MaxNumToReturn = 0; // 0 means no limit

                    qryInput[0].Entries = new String[] { "User", };

                    qryInput[0].Values = new String[1];

                    qryInput[0].Values[0] = userID;


                    Teamcenter.Services.Strong.Query._2007_09.SavedQuery.SavedQueriesResponse executeQry = TcAdaptor.savedQryServices.ExecuteSavedQueries(qryInput);

                    if (executeQry.ArrayOfResults.Length == 0)
                    {
                        return groups;
                    }

                    Teamcenter.Services.Strong.Query._2007_09.SavedQuery.QueryResults qryResult = executeQry.ArrayOfResults[0];


                    //  ServiceData sData = Tc_Services.dmService.LoadObjects(qryResult.ObjectUIDS);

                    ServiceData sData = TcAdaptor.dmService.LoadObjects(qryResult.ObjectUIDS);

                    for (int k = 0; k < sData.sizeOfPlainObjects(); k++)
                    {
                        GroupMember grpMem = (GroupMember)sData.GetPlainObject(k);

                        Group _grp = (Group)grpMem.The_group;

                        ServiceData _sData = TcAdaptor.dmService.LoadObjects(new string[] { _grp.Uid });

                        for (int n = 0; n < _sData.sizeOfPlainObjects(); n++)
                        {
                            Group grp = (Group)_sData.GetPlainObject(n);

                            groups.Add(grp);
                        }

                    }

                    return groups;
                }
                else
                {
                    Utility.Log("qry_EINT_group_members... query not found in Teamcenter.", logfilepath);
                    return groups;
                }


            }
            catch (Exception e)
            {
               Utility.Log("Exception in get_ItemRev_ModelObj" + e,logfilepath);
                return groups;
            }

        }
    }
}