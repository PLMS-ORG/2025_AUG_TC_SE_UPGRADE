using DemoAddInTC.model;
using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC.CustomView
{
    class MyCustomDialog4View : MyCustomDialog4
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

        public static void selectAllNodesWithText(TreeView treeView, String occurrenceName, bool checkStatus, String logFilePath)
        {
            //TraverseNodes(treeNode, occurrenceName, checkStatus);            
            Utlity.Log("occurrenceName " + occurrenceName, logFilePath);
            TraverseTree(treeView, occurrenceName, checkStatus, logFilePath);
        }

        //// OccurenceName Match in the TreeView
        private static void TraverseTree(TreeView treeView, String occurrenceName, bool checkStatus, String logFilePath)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            if (nodes == null) return;
            if (nodes.Count == 0) return;
            foreach (TreeNode n in nodes)
            {
                if (n.Text != null && n.Text.Equals("") == false)
                {

                    Utlity.Log("n.Text " + n.Text, logFilePath);
                    if (n.Text.Equals(occurrenceName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        n.Checked = checkStatus;
                    }
                }
                TreeNodeCollection childNodes = n.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(n, occurrenceName, checkStatus, logFilePath);
                }
            }

        }

        //// OccurenceName Match in the TreeView
        public static void TraverseNodes(TreeNode treeNode, String occurrenceName, bool checkStatus, String logFilePath)
        {
            if (treeNode.Nodes == null) return;
            if (treeNode.Nodes.Count == 0) return;
            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                if (tn.Text != null && tn.Text.Equals("") == false)
                {

                    Utlity.Log("tn.Text " + tn.Text, logFilePath);
                    if (tn.Text.Equals(occurrenceName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        tn.Checked = checkStatus;
                    }
                }

                TreeNodeCollection childNodes = tn.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(tn, occurrenceName, checkStatus, logFilePath);
                }
            }
        }

        // SELECT ALL / UNSELECT ALL
        public static void TraverseNodes(TreeNode treeNode, bool checkStatus)
        {
            if (treeNode.Nodes == null) return;
            if (treeNode.Nodes.Count == 0) return;
            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                if (tn.Text != null && tn.Text.Equals("") == false)
                {
                    tn.Checked = checkStatus;
                }

                TreeNodeCollection childNodes = tn.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(tn, checkStatus);
                }
            }
        }

        // -------START// 
        // Manage Components -- 
        public static Dictionary<String, bool> getSelectedComponentsFromHierarchialTreeView(TreeView treeView1)
        {
            Dictionary<String, bool> partEnabledDictionary = MyCustomDialog4.getPartEnabledOrNotDictionary();
            TraverseTree(treeView1, partEnabledDictionary);
            return partEnabledDictionary;
        }

        // Manage Components --  
        // Call the procedure using the TreeView.
        private static void TraverseTree(TreeView treeView, Dictionary<String, bool> partEnabledDictionary)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            foreach (TreeNode n in nodes)
            {
                String occName = n.Text;
                String nodeText = "";
                if (occName != null && occName.Equals("") == false)
                {
                    String[] occNameArray = occName.Split(';');
                    if (occNameArray != null && occNameArray.Length != 0)
                    {
                        nodeText = occNameArray[0];
                        occName = nodeText;
                    }
                }
                if (n.Checked == true)
                {
                    if (partEnabledDictionary.ContainsKey(occName) == false)
                    {
                        partEnabledDictionary.Add(occName, true);
                    }
                    else
                    {
                        partEnabledDictionary[occName] = true;
                    }
                    //MessageBox.Show(occName + ":::" + "TRUE");
                }
                else
                {
                    if (partEnabledDictionary.ContainsKey(occName) == false)
                    {
                        partEnabledDictionary.Add(occName, false);
                    }
                    else
                    {
                        partEnabledDictionary[occName] = false;
                    }
                    //MessageBox.Show(occName + ":::" + "FALSE");
                }
                TreeNodeCollection childNodes = n.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(n, partEnabledDictionary);
                }
            }
        }

        // Manage Components
        public static void TraverseNodes(TreeNode treeNode, Dictionary<String, bool> partEnabledDictionary)
        {
            if (treeNode.Nodes == null) return;
            if (treeNode.Nodes.Count == 0) return;

            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                String occName = tn.Text;
                if (occName != null && occName.Equals("") == false)
                {
                    String[] occNameArray = occName.Split(';');
                    if (occNameArray != null && occNameArray.Length != 0)
                    {
                        String nodeText = occNameArray[0];
                        occName = nodeText;
                    }
                }
                if (tn.Checked == true)
                {
                    if (partEnabledDictionary.ContainsKey(occName) == false)
                    {
                        partEnabledDictionary.Add(occName, true);
                    }
                    else
                    {
                        partEnabledDictionary[occName] = true;
                    }
                    
                    //MessageBox.Show(occName + ":::" + "TRUE");
                }
                else
                {
                    if (partEnabledDictionary.ContainsKey(occName) == false)
                    {
                        partEnabledDictionary.Add(occName, false);
                    }
                    else
                    {
                        partEnabledDictionary[occName] = false;
                    }
                    //MessageBox.Show(occName + ":::" + "FALSE");
                }

                TreeNodeCollection childNodes = tn.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(tn, partEnabledDictionary);
                }
            }
        }
        // -------END // 

        // Get Enabled/Disabled State from partEnabledDictionary & SET to NODE.
        public static void TraverseTree(TreeView treeView, Dictionary<String, bool> partEnablementDictionary, String logFilePath)
        {
            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                return;
            }
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            if (nodes == null) return;
            if (nodes.Count == 0) return;
            foreach (TreeNode n in nodes)
            {
                if (n.Text != null && n.Text.Equals("") == false)
                {

                    Utlity.Log("n.Text " + n.Text, logFilePath);
                    bool value = false;
                    bool Success = partEnablementDictionary.TryGetValue(n.Text, out value);
                    if (Success == true)
                    {
                        n.Checked = value;
                    }
                    else
                    {
                        n.Checked = value;
                    }
                }
                TreeNodeCollection childNodes = n.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(n, partEnablementDictionary, logFilePath);
                }
            }

        }

        public static void TraverseNodes(TreeNode treeNode, Dictionary<String, bool> partEnablementDictionary, String logFilePath)
        {
            if (treeNode.Nodes == null) return;
            if (treeNode.Nodes.Count == 0) return;
            // Print each node recursively.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                if (tn.Text != null && tn.Text.Equals("") == false)
                {

                    Utlity.Log("tn.Text " + tn.Text, logFilePath);
                    String tn_Text = tn.Text;
                    String[] TextArray = tn_Text.Split(';');
                    if (TextArray == null || TextArray.Length == 0) continue;
                    String item_Id = TextArray[0];
                    bool value = false;
                    bool Success = partEnablementDictionary.TryGetValue(item_Id, out value);
                    if (Success == true)
                    {
                        tn.Checked = value;
                    }
                    else
                    {
                        tn.Checked = value;
                    }
                }

                TreeNodeCollection childNodes = tn.Nodes;
                if (childNodes == null) continue;
                if (childNodes.Count != 0)
                {
                    TraverseNodes(tn, partEnablementDictionary, logFilePath);
                }
            }
        }
    }
}
