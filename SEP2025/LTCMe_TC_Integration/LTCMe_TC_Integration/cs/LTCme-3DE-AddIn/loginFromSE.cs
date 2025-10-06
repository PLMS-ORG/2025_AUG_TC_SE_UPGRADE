using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DemoAddInTC.se;
using DemoAddInTC.utils;
using get_group_role_member;

namespace DemoAddInTC
{
    public partial class loginFromSE : Form
    {
        public loginFromSE()
        {
            string PropertyFile = null;
            PropertyFile = Utlity.getPropertyFilePath();
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                MessageBox.Show("Could Not Find Property File " + PropertyFile);
                return;
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                MessageBox.Show("Cannot Parse Property File to Get Templates Folder: " + PropertyFile);
                return;
            }
            SeecURLs = Props.get("SEEC_URL");

            InitializeComponent();
        }

        public static void ParsePropertyFile()
        {
            string PropertyFile = null;
            PropertyFile = Utlity.getPropertyFilePath();
            if (PropertyFile == null || PropertyFile.Equals("") == true)
            {
                MessageBox.Show("Could Not Find Property File " + PropertyFile);
                return;
            }
            utils.Property Props = new utils.Property(PropertyFile);
            if (Props == null)
            {
                MessageBox.Show("Cannot Parse Property File to Get Templates Folder: " + PropertyFile);
                return;
            }
            SeecURLs = Props.get("SEEC_URL");
            string[] NameAndURL = SeecURLs.Split('=');
            if (NameAndURL != null && NameAndURL.Length == 2)
            {
                URL= NameAndURL[1];
            }
        }

        public static string userName = null, password = null, group = null, role = null, URL = null, SeecURLs = null;
        public static bool loggedInthroughUtility = false;

        private void okButton_Click(object sender, EventArgs e)
        {
            if (userNameTextBox.Text.Trim().Equals("") == true)
            {
                MessageBox.Show("Username cannot be blank");
                return;
            }

            if (passwordTextBox.Text.Trim().Equals("") == true)
            {
                MessageBox.Show("Password cannot be blank");
                return;
            }

            if (URLComboBox.SelectedItem == null)
            {
                MessageBox.Show("URL cannot be blank");
                return;
            }

            userName = userNameTextBox.Text;
            password = passwordTextBox.Text;
            NameAndURLDictionary.TryGetValue(URLComboBox.SelectedItem.ToString(), out URL);

            role = roleComboBox.Text;
            group = groupComboBox.Text;

            this.okButton.Enabled = false;
            this.DialogResult = DialogResult.OK;

        }

        //============================================================================================
        //added by pragati Date: 05-04-2024
        // when user selects any group from groupcombobox drop down,
        // then role combo box items should get updated as per selected group for corresponding user
        private void groupComboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
          //  MessageBox.Show("call to groupComboBox_SelectedIndexChanged_1");
            List<string> rolename;
            if (groupVsRoleNames.TryGetValue(groupComboBox.SelectedItem.ToString(), out rolename))
                SetRoleComboBoxDropList(rolename);
        }

        

        private void populateGroupAndRole(string group, string role)
        {
            string logfileDir = Utlity.CreateLogDirectory();
            String groupRoleLog = System.IO.Path.Combine(logfileDir, "LTC_FetchingGroupRoleFromTC.txt");
            Utility.Log("populateGroupAndRole", groupRoleLog);

            NameAndURLDictionary = new Dictionary<string, string>();

            //groupComboBox.Text = group;
            //Utility.Log("populateGroupAndRole: Group: "+ group, groupRoleLog);
            //roleComboBox.Text = role;
            //Utility.Log("populateGroupAndRole: role: " + role, groupRoleLog);

            string[] NameAndURL = SeecURLs.Split('~');
            foreach (string s in NameAndURL)
            {
                string[] splitString = s.Split('=');
                URLComboBox.Items.Add(splitString.First());
                NameAndURLDictionary.Add(splitString.First(), splitString.Last());
            }
            if (URLComboBox.Items.Count >= 1)
                URLComboBox.SelectedIndex = 0;

            //added by pragati Date: 05-04-2024
            //when user again try to login as designer in same session
            //use same username and password
            //user allowed to change group and role only 
            //password and username text box are set as read only
            Utility.Log("populateGroupAndRole: userName: " + userNameTextBox.Text, groupRoleLog);
            Utility.Log("populateGroupAndRole: password: " + passwordTextBox.Text, groupRoleLog);
            userName = userNameTextBox.Text;
            password = passwordTextBox.Text;
            if (!string.IsNullOrWhiteSpace(userNameTextBox.Text))
            {
                //userNameTextBox.Text = userName;
                userNameTextBox.ReadOnly = true;

                if (passwordTextBox.Text != string.Empty)
                {
                    //passwordTextBox.Text = password;
                    passwordTextBox.ReadOnly = true;
                }

                //if group and corresponding group names are not fetched in current session then collect it 
                //it should be fetched only once in current session so groupVsRolenames is kept as static variable
                groupVsRoleNames.Clear();
                Utility.Log("populateGroupAndRole: groupVsRoleNames Count : " + groupVsRoleNames.Count, groupRoleLog);
                //if (groupVsRoleNames.Count == 0)
                {
                    //string logfileDir = Utlity.CreateLogDirectory();
                    //String groupRoleLog = System.IO.Path.Combine(logfileDir, "LTC_FetchingGroupRoleFromTC.txt");
                    Utility.Log("populateGroupAndRole: login to TC using empty group and role: ", groupRoleLog);
                    group = "";
                    role = "";
                    Utility.Log("populateGroupAndRole: group : "+ group, groupRoleLog);
                    Utility.Log("populateGroupAndRole: role :" + role, groupRoleLog);
                    TcAdaptor.login(userName, password, group, role, groupRoleLog);
                    Utility.Log("populateGroupAndRole: get_User_Groups_And_Roles: ", groupRoleLog);
                    GroupRoleUserUtil.get_User_Groups_And_Roles(userName, groupRoleLog, ref groupVsRoleNames);
                    Utility.Log("populateGroupAndRole: logout: ", groupRoleLog);
                    TcAdaptor.logout(groupRoleLog);
                    if (groupVsRoleNames.Count == 0)
                    {
                        MessageBox.Show("Groups and roles information not fetched for username :" + userName);
                        this.DialogResult = DialogResult.Cancel;
                        return;
                    }
                }
                Utility.Log("populateGroupAndRole: SetGroupComboBoxDropList: ", groupRoleLog);
                SetGroupComboBoxDropList(groupVsRoleNames.Keys.ToList());
            }

        }

        private void passwordTextBox_Leave(object sender, EventArgs e)
        {
            if (userNameTextBox.Text == null || userNameTextBox.Text == "" ||
                userNameTextBox.Text.Equals("") == true)
            {
                // do nothing
                return;
            }

            populateGroupAndRole(group, role);

        }

        private void userNameTextBox_Leave(object sender, EventArgs e)
        {
            if (passwordTextBox.Text == null || passwordTextBox.Text == "" ||
                passwordTextBox.Text.Equals("") == true)
            {
                // do nothing
                return;
            }
            populateGroupAndRole(group, role);

        }

        //==============================================================================================

        Dictionary<string, string> NameAndURLDictionary = new Dictionary<string, string>();

        //dictionary of group vs its corresponding role for user
        static Dictionary<string,List<string>>groupVsRoleNames = new Dictionary<string,List<string>>();
        private void loginFromSE_Load(object sender, EventArgs e)
        {
            //NameAndURLDictionary = new Dictionary<string, string>();

            //groupComboBox.Text = group;

            //roleComboBox.Text = role;

            //string[] NameAndURL = SeecURLs.Split('~');
            //foreach (string s in NameAndURL)
            //{
            //    string[] splitString = s.Split('=');
            //    URLComboBox.Items.Add(splitString.First());
            //    NameAndURLDictionary.Add(splitString.First(), splitString.Last());
            //}
            //if (URLComboBox.Items.Count >= 1)
            //    URLComboBox.SelectedIndex = 0;

            ////added by pragati Date: 05-04-2024
            ////when user again try to login as designer in same session
            ////use same username and password
            ////user allowed to change group and role only 
            ////password and username text box are set as read only
            //if (!string.IsNullOrWhiteSpace(userName))
            //{
            //    userNameTextBox.Text = userName;
            //    userNameTextBox.ReadOnly = true;

            //    if (password != string.Empty)
            //    {
            //        passwordTextBox.Text = password;
            //        passwordTextBox.ReadOnly = true;
            //    }

            //    //if group and corresponding group names are not fetched in current session then collect it 
            //    //it should be fetched only once in current session so groupVsRolenames is kept as static variable
            //    if (groupVsRoleNames.Count == 0)
            //    {
            //        string logfileDir = Utlity.CreateLogDirectory();
            //        String logfilepath = System.IO.Path.Combine(logfileDir, "FetchingGroupRoleFromTC.txt");

            //        TcAdaptor.login(userName, password, group, role, logfilepath);
            //        GroupRoleUserUtil.get_User_Groups_And_Roles(userName, logfilepath, ref groupVsRoleNames);
            //        TcAdaptor.logout(logfilepath);
            //        if (groupVsRoleNames.Count == 0)
            //        {
            //            MessageBox.Show("Groups and roles information not fetched for username :" + userName);
            //            this.DialogResult = DialogResult.Cancel;
            //            return;
            //        }
            //    }
            //    SetGroupComboBoxDropList(groupVsRoleNames.Keys.ToList());
            //}
        }


        //============================================================================================================
        //add items to group combo box
        public void SetGroupComboBoxDropList(List<string>groups)
        {
            string logfileDir = Utlity.CreateLogDirectory();
            String groupRoleLog = System.IO.Path.Combine(logfileDir, "LTC_FetchingGroupRoleFromTC.txt");
            Utility.Log("SetGroupComboBoxDropList", groupRoleLog);

            // Create a HashSet to track unique strings with case sensitivity
            HashSet<string> uniqueStrings = new HashSet<string>(groups);

            // Clear the original list and add only the unique values back
            List<string> groups_new = new List<string>(uniqueStrings);

            // Output the list after removing duplicates
            foreach (string s in groups_new)
            {
                Utility.Log(s, groupRoleLog);
            }

            groupComboBox.Items.Clear();
            foreach (string s in groups_new)
             groupComboBox.Items.Add(s);

            if (groupComboBox.Items.Count >= 1)
                groupComboBox.SelectedIndex = 0;
        }


        //=============================================================================================================
        //add items to role combo box 
        public void SetRoleComboBoxDropList(List<string> roles)
        {
            string logfileDir = Utlity.CreateLogDirectory();
            String groupRoleLog = System.IO.Path.Combine(logfileDir, "LTC_FetchingGroupRoleFromTC.txt");
            Utility.Log("SetRoleComboBoxDropList", groupRoleLog);

            // Create a HashSet to track unique strings with case sensitivity
            HashSet<string> uniqueStrings = new HashSet<string>(roles);

            // Clear the original list and add only the unique values back
            List<string> roles_new = new List<string>(uniqueStrings);

            // Output the list after removing duplicates
            foreach (string s in roles_new)
            {
                Utility.Log(s, groupRoleLog);

            }
            roleComboBox.Items.Clear();
            foreach (string s in roles_new)
               roleComboBox.Items.Add(s);

            if (groupComboBox.Items.Count >= 1)
                roleComboBox.SelectedIndex = 0;
        }

        //=============================================================================================================

    }
}
