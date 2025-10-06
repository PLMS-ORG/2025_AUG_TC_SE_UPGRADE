using DemoAddInTC.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC.CustomView
{
    public class MyCustomDialog3View : MyCustomDialog3
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

        public static void selectAllNodesWithText(TreeView treeView, String occurrenceName, bool checkStatus,String logFilePath)
        {
            //TraverseNodes(treeNode, occurrenceName, checkStatus);            
            Utlity.Log("occurrenceName " + occurrenceName, logFilePath);
            TraverseTree(treeView, occurrenceName, checkStatus,logFilePath);
        }

        //// Call the procedure using the TreeView.
        private static void TraverseTree(TreeView treeView, String occurrenceName, bool checkStatus,String logFilePath)
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
                    TraverseNodes(n, occurrenceName, checkStatus,logFilePath);
                }
            }

        }

        private static void TraverseNodes(TreeNode treeNode, String occurrenceName, bool checkStatus, String logFilePath)
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
                    TraverseNodes(tn, occurrenceName, checkStatus,logFilePath);
                }
            }
        }


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

//-------------------------------------//
        // Manage Components -- 
        public static Dictionary<String, bool> getSelectedComponentsFromHierarchialTreeView(TreeView treeView1)
        {
            Dictionary<String,bool> partEnabledDictionary = MyCustomDialog3.getPartEnabledOrNotDictionary();
            TraverseTree(treeView1, partEnabledDictionary);
            return partEnabledDictionary;
        }

        // Manage Components
        private static void TraverseTree(TreeView treeView, Dictionary<String, bool> partEnabledDictionary)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            //MessageBox.Show("Nodes Count:" + nodes.Count);
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
                    if (partEnabledDictionary.ContainsKey(occName) == false) // 
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
                    //MessageBox.Show("childNodes Count:" + childNodes.Count);
                    TraverseNodes(n, partEnabledDictionary);
                }
            }
        }

        // Manage Components
        private static void TraverseNodes(TreeNode treeNode, Dictionary<String, bool> partEnabledDictionary)
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
    }
    //-------------------------------------//
}
