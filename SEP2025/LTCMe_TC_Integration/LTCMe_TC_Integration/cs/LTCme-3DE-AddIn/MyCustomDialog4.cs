/* 3DE International FZE, 2018
 *  Murali - 02-OCT-18 - Changes for SuppressionEnabled Property
 */

using DemoAddInTC.controller;
using DemoAddInTC.CustomView;
using DemoAddInTC.FormulaEvaluator;
using DemoAddInTC.model;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.utils;
using DgvFilterPopup;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using DemoAddInTC.se;
using System.IO;

namespace DemoAddInTC
{
    public partial class MyCustomDialog4 : Form
    {
        Subro.Controls.DataGridViewGrouper grouper = null;
        Subro.Controls.DataGridViewGrouper grouper_ManageFeatures = null;
        static Dictionary<String, bool> partEnabledDictionary = new Dictionary<string, bool>();
        static List<Variable> allVarList = new List<Variable>();
        static List<FeatureLine> allFeatureList = new List<FeatureLine>();

        public static Dictionary<String, bool> getPartEnabledOrNotDictionary()
        {
            return partEnabledDictionary;
        }

        public MyCustomDialog4()
        {
            InitializeComponent();
            try
            {
                //FillComponentList();
                FillComponentListHierarchy();
            }
            catch (Exception ex)
            {
                MessageBox.Show("FillComponentList Exception: " + ex.Message);
                return;
            }
            try
            {
                getSelectedComponentsFromTreeView();
            }
            catch (Exception ex)
            {
                MessageBox.Show("getSelectedComponentsFromTreeView Exception: " + ex.Message);
                return;
            }
            try
            {
                populateDataFromExcel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("populateDataFromExcel Exception: " + ex.Message);
                return;
            }
            try
            {
                applyGroup();
            }
            catch (Exception ex)
            {
                MessageBox.Show("applyGroup Exception: " + ex.Message);
              return;
            }
        }

        private void FillComponentListHierarchy()
        {
            List<String> bomLinesList = null;
            Dictionary<String, bool> partEnableDict = ExcelData.getPartEnablementDictionary();
            //MessageBox.Show(partEnableDict.Keys.Count.ToString());
            bomLinesList = ExcelData.getOcurrenceList();
            if (bomLinesList == null || bomLinesList.Count == 0 || partEnableDict == null || partEnableDict.Count == 0)
            {
                return;
            }
            if (treeView1 == null)
            {
                MessageBox.Show("Unable to Initialize TreeView");
                return;
            }

            TreeNode tNode;
            try
            {
                treeView1.Nodes.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;

            }
            Dictionary<String, String> SKYnameDictionary = SolidEdgeHighLighter.getSKYnameDictionary();
            if (SKYnameDictionary.Keys.Count == 0) 
            {
                MessageBox.Show("SKYnameDictionary is Empty...ItemName is missing for Some/All items of this Assembly. update the ItemName...");
                return; 
            }

            String SKYname = "";
            if (SKYnameDictionary.ContainsKey(bomLinesList[0]) == true)
            {
                SKYnameDictionary.TryGetValue(bomLinesList[0], out SKYname);
            }
            if (SKYname == null || SKYname.Equals("") == true)
            {
                MessageBox.Show("ItemName is Empty for..." + bomLinesList[0] + "..update the ItemName...");
                return;
            }
            String item_name = "";
            if (SKYname.Contains(".") == true)
            {
                String[] SKYnameArray = SKYname.Split('.');
                if (SKYnameArray == null || SKYnameArray.Length == 0) return;
                item_name = SKYnameArray[0];
            }

            // only for Top Node
            tNode = treeView1.Nodes.Add(bomLinesList[0], bomLinesList[0] + ";" + item_name, 1);
            {
                tNode.ImageIndex = 1;
                tNode.SelectedImageIndex = 1;
            }
            bool value = false;
            partEnableDict.TryGetValue(bomLinesList[0], out value);
            if (value == true)
            {
                treeView1.Nodes[0].Checked = true;
            }

            String LogStageDir = Utlity.CreateLogDirectory();

            //String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(SolidEdgeData1.getAssemblyFileName()) + "_" + "LISTHIERARCHY" + ".txt");
            String logFilePath = "";
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            SolidEgeData2.traverseAssembly1(SolidEdgeData1.getAssemblyFileName(), logFilePath, tNode);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            try
            {
                MyCustomDialog4View.TraverseNodes(tNode, partEnableDict,logFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            TreeViewSettings();

        }

        private void PopulateFeatureDataFromExcel()
        {
            allFeatureList.Clear();

            Dictionary<String, List<FeatureLine>> featureDictionary = ExcelReadFeatures.getFeatureDictionary();

            if (featureDictionary == null || featureDictionary.Count == 0)
            {
                //MessageBox.Show("No Features in Assembly to Show");
                return;
            }

            if (partEnabledDictionary == null || partEnabledDictionary.Count == 0)
            {
                MessageBox.Show("Data Missing to Show the UI: partEnabledDictionary");
                return;
            }
            int childCnt = partEnabledDictionary.Count;
            foreach (String s in partEnabledDictionary.Keys)
            {
                String occName = s;
                bool value = false;
                partEnabledDictionary.TryGetValue(s, out value);
                if (value == true)
                {
                    List<FeatureLine> featuresList = null;
                    bool Success = featureDictionary.TryGetValue(occName, out featuresList);
                    //MessageBox.Show(occName + ":::" + featuresList.Count.ToString());
                    if (Success == true)
                    {
                        allFeatureList.AddRange(featuresList);
                    }
                }
            }

            //allFeatureList = SolidEdgeReadFeature.getFeatureLines();
            if (allFeatureList == null || allFeatureList.Count == 0)
            {
                //MessageBox.Show("No Features in Assembly to Show");
                return;
            }

            List<FeatureLine> newAllFsList = PopulateSKYname(allFeatureList);
            if (newAllFsList == null || newAllFsList.Count == 0)
            {
                MessageBox.Show("No Features to show..");
                return;

            }

            DataTable table = null;
            table = Utlity.ConvertToDataTable(newAllFsList);
            if (table == null)
            {
                MessageBox.Show("Unable to Convert Variables to Table");
                return;
            }

            //MessageBox.Show(table.Rows.Count.ToString());            
            this.dataGridView2.DataSource = table;


            try
            {
                applyGroup_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Apply Group to  Feature Table");
                return;

            }


            try
            {
                ApplyAdditionalSettings_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Apply Settings to  Feature Table");
                return;

            }

            try
            {
                hideColumns_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to hide Columns to  Feature Table");
                return;

            }

        }

        private void applyGroup_ManageFeatures()
        {
            try
            {
                if (grouper_ManageFeatures == null)
                {
                    grouper_ManageFeatures = new Subro.Controls.DataGridViewGrouper(dataGridView2);
                }
                grouper_ManageFeatures.SetGroupOn<FeatureLine>(a => a.SKYname);
                grouper_ManageFeatures.CollapseAll();
                grouper_ManageFeatures.DisplayGroup += grouper_ManageFeatures_DisplayGroup;
                grouper_ManageFeatures.GroupingChanged += grouper_ManageFeatures_GroupingChanged;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void grouper_ManageFeatures_GroupingChanged(object sender, EventArgs e)
        {
            //grouper_ManageFeatures.RemoveGrouping();
            //MessageBox.Show("Inside GroupingChanged");

            //MessageBox.Show(sender.GetType().ToString());

            grouper_ManageFeatures.SetGroupOn<FeatureLine>(a => a.SKYname);
            grouper_ManageFeatures.CollapseAll();

        }

        private void grouper_GroupingChanged(object sender, EventArgs e)
        {
            grouper.SetGroupOn<Variable>(t => t.Skyname);
            grouper.CollapseAll();

        }

        private void hideColumns_ManageFeatures()
        {
            this.dataGridView2.Columns["SystemName"].Visible = false;
            //this.dataGridView2.Columns["PartName"].Visible = false;
            // 02 OCT - based on Request from LTC, Added New Column
            this.dataGridView2.Columns["IsFeatureEnabled"].Visible = false;
            this.dataGridView2.Columns["SKYname"].Visible = false;
        }

        private void ApplyAdditionalSettings_ManageFeatures()
        {
            this.dataGridView2.Columns["FeatureName"].ReadOnly = true;

            // 02 OCT - based on Request from LTC, Added New Column
            this.dataGridView2.Columns["SuppressionEnabled"].ReadOnly = false;
            this.dataGridView2.Columns["IsFeatureEnabled"].ReadOnly = true;
            this.dataGridView2.Columns["EdgeBarName"].ReadOnly = true;
            this.dataGridView2.Columns["SKYname"].ReadOnly = true;

            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.Azure;
        }

        //optionally, you can customize the grouping display by subscribing to the DisplayGroup event
        void grouper_ManageFeatures_DisplayGroup(object sender, Subro.Controls.GroupDisplayEventArgs e)
        {
            e.BackColor = (e.Group.GroupIndex % 2) == 0 ? Color.LightBlue : Color.White;
            //e.BackColor = Color.LightGray;
            //e.Header = "[" + e.Header + "], grp: " + e.Group.GroupIndex;
            e.Header = "";
            e.DisplayValue = e.DisplayValue;
            e.Summary = "(" + e.Group.Count + " )";
            e.ForeColor = Color.Black;
        }

        private void applyGroup()
        {
            try
            {

                grouper = new Subro.Controls.DataGridViewGrouper(dataGridView1);
                grouper.SetGroupOn<Variable>(t => t.Skyname);
                grouper.CollapseAll();
                grouper.DisplayGroup += grouper_DisplayGroup;
                grouper.GroupingChanged += grouper_GroupingChanged;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void grouper_GroupingChanged(object sender, EventArgs e)
        //{
        //    grouper.SetGroupOn<Variable>(t => t.PartName);
        //    grouper.CollapseAll();

        //}

        private void getSelectedComponentsFromTreeView()
        {
            partEnabledDictionary.Clear();
            //int childCnt = treeView1.Nodes[0].Nodes.Count;
            //for (int i = 0; i < childCnt; i++)
            //{
            //    String occName = treeView1.Nodes[0].Nodes[i].Text;
            //    if (treeView1.Nodes[0].Nodes[i].Checked == true)
            //    {
            //        partEnabledDictionary.Add(occName, true);
            //        //MessageBox.Show(occName + ":::" + "TRUE");
            //    }
            //    else
            //    {
            //        partEnabledDictionary.Add(occName, false);
            //        //MessageBox.Show(occName + ":::" + "FALSE");
            //    }

            //}

            partEnabledDictionary = MyCustomDialog4View.getSelectedComponentsFromHierarchialTreeView(treeView1);  

        }

        private void FillComponentList()
        {
            List<String> bomLinesList = null;
            Dictionary<String,bool> partEnableDict = ExcelData.getPartEnablementDictionary();
            //MessageBox.Show(partEnableDict.Keys.Count.ToString());
            bomLinesList = ExcelData.getOcurrenceList();
            if (bomLinesList == null || bomLinesList.Count == 0 || partEnableDict==null ||  partEnableDict.Count == 0)
            {
                return;
            }
            if (treeView1 == null)
            {
                MessageBox.Show("Unable to Initialize TreeView");
                return;
            }
            TreeNode tNode;
            try
            {
                treeView1.Nodes.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;

            }
            
            tNode = treeView1.Nodes.Add(bomLinesList[0], bomLinesList[0], 1);
            bool value = false;
            partEnableDict.TryGetValue(bomLinesList[0], out value);
            if (value == true)
            {
                treeView1.Nodes[0].Checked = true;
            }

            //MessageBox.Show(bomLinesList.Count.ToString());
            for (int i = 1; i < bomLinesList.Count; i++)
            {
                if (bomLinesList[i] == null || bomLinesList[i].Equals("") == true)
                {
                    continue;
                }
                //MessageBox.Show(bomLinesList[i]);
                if (bomLinesList[i].EndsWith(".par", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //treeView1.Nodes[0].Nodes.Add(bomLinesList[i]);
                    treeView1.Nodes[0].Nodes.Add(bomLinesList[i], bomLinesList[i], 0);
                }
                else if (bomLinesList[i].EndsWith(".psm", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //treeView1.Nodes[0].Nodes.Add(bomLinesList[i]);
                    treeView1.Nodes[0].Nodes.Add(bomLinesList[i], bomLinesList[i], 1);
                }
                else if (bomLinesList[i].EndsWith(".asm", StringComparison.OrdinalIgnoreCase) == true)
                {
                    //treeView1.Nodes[0].Nodes.Add(bomLinesList[i]);
                    treeView1.Nodes[0].Nodes.Add(bomLinesList[i], bomLinesList[i], 1);
                }

                value = false;
                partEnableDict.TryGetValue(bomLinesList[i], out value);
                if (value == true)
                {
                    // TreeView nodes start with Index 1 Only
                    treeView1.Nodes[0].Nodes[i-1].Checked = true;
                }

                //if (i % 2 == 0)
                //{
                //    treeView1.Nodes[0].Nodes[i - 1].BackColor = Color.LightBlue;
                //}
                //else
                //{
                //    //treeView1.Nodes[0].Nodes[i - 1].BackColor = Color.LightBlue;
                //}
            }

            TreeViewSettings();
        }

        private void TreeViewSettings()
        {
            // Set the BorderStyle property to none, the BackColor property to  
            // the form's backcolor, and the Scrollable property to false.  
            // This allows the TreeView to blend in form.

            this.treeView1.BorderStyle = BorderStyle.None;
            this.treeView1.BackColor = this.BackColor;
            this.treeView1.Scrollable = false;

            // Set the HideSelection property to false to keep the 
            // selection highlighted when the user leaves the control. 
            // This helps it blend with form.
            this.treeView1.HideSelection = false;

            // Set the ShowRootLines and ShowLines properties to false to 
            // give the TreeView a list-like appearance.
            this.treeView1.ShowRootLines = true;
            this.treeView1.ShowLines = true;
            this.treeView1.Scrollable = true;
        }

        private void ClubUIAndDataFromSE()
        {
            DataTable ds = (DataTable)grouper.GroupingSource.DataSource;
            List<Variable> UIvariablesList = null;
            // logFilePath is Empty -- No Need to LOG
            if (ds != null)
            {
                UIvariablesList = ConvertDataTableToList.ConvertDataTableToVariablesList(ds, "");
            }
            Dictionary<String, List<Variable>> UIVarDictionary = null;
            if (UIvariablesList != null && UIvariablesList.Count != 0)
            {
                UIVarDictionary = Utlity.BuildVariableDictionary(UIvariablesList, "");
                //MessageBox.Show("UIVarDictionary.Count" + UIVarDictionary.Count.ToString());
            }
            else
            {
                UIVarDictionary = null;
            }

            allVarList.Clear();

            Dictionary<String, List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
            if (variableDictionary.Count == 0 || variableDictionary == null)
            {
                MessageBox.Show("No Variables in SolidEdge to Show");
                return;
            }
            if (partEnabledDictionary == null || partEnabledDictionary.Count == 0)
            {
                MessageBox.Show("Data Missing to Show the UI: partEnabledDictionary");
                return;
            }
            int childCnt = partEnabledDictionary.Count;
            foreach (String s in partEnabledDictionary.Keys)
            {
                String occName = s;
                bool value = false;
                partEnabledDictionary.TryGetValue(s, out value);
                if (value == true)
                {
                    List<Variable> variablesList = null;

                    if (UIVarDictionary != null && UIVarDictionary.Count != 0)
                    {
                        if (UIVarDictionary.ContainsKey(s) == true)
                        {
                            bool Success = UIVarDictionary.TryGetValue(occName, out variablesList);
                            //MessageBox.Show(occName + ":::" + variablesList.Count.ToString());
                            if (Success == true)
                            {
                                if (variablesList != null & variablesList.Count != 0)
                                {
                                    allVarList.AddRange(variablesList);
                                }
                            }
                        }
                        else
                        {
                            bool Success = variableDictionary.TryGetValue(occName, out variablesList);
                            //MessageBox.Show(occName + ":::" + variablesList.Count.ToString());
                            if (Success == true)
                            {
                                if (variablesList != null & variablesList.Count != 0)
                                {
                                    allVarList.AddRange(variablesList);
                                }
                            }
                        }
                    }
                    else
                    {
                        bool Success = variableDictionary.TryGetValue(occName, out variablesList);
                        //MessageBox.Show(occName + ":::" + variablesList.Count.ToString());
                        if (Success == true)
                        {
                            if (variablesList != null & variablesList.Count != 0)
                            {
                                allVarList.AddRange(variablesList);
                            }
                        }
                    }

                }
            }

            List<Variable> newAllVarsList = PopulateSKYname(allVarList);
            if (newAllVarsList == null || newAllVarsList.Count == 0)
            {
                MessageBox.Show("No variables to show..");
                return;

            }

            DataTable table = null;
            table = Utlity.ConvertToDataTable(newAllVarsList);
            if (table == null)
            {
                MessageBox.Show("Unable to Convert Variables to Table");
                return;
            }

            //MessageBox.Show(table.Rows.Count.ToString());            
            this.dataGridView1.DataSource = table;
            try
            {
                hideColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in Hiding Columns: " + ex.Message);
                return;
            }
            try
            {
                ApplyAdditionalSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in Applying Addition Settings: " + ex.Message);
                return;
            }
            try
            {
                setColumnsSettingsInDGV();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in Setting Columns in DGV : " + ex.Message);
                return;
            }

            try
            {
                AppyHeaderNames();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in Setting Column Names in DGV : " + ex.Message);
                return;
            }

            if (UIVarDictionary != null)
                UIVarDictionary.Clear();
            if (UIvariablesList != null)
                UIvariablesList.Clear();
            if (allVarList != null)
                allVarList.Clear();
        }

        private void ClubUIAndFeatureDataFromSE()
        {
            if (allFeatureList != null)
                allFeatureList.Clear();

            Dictionary<String, List<FeatureLine>> featureDictionary = null;
            try
            {
                featureDictionary = ExcelReadFeatures.getFeatureDictionary();
                //File.AppendAllLines("C:\\Users\\3de.AE-LTC\\3DE_LTC\\log.txt", new string[] { "featureDictionary.Count " + featureDictionary.Count.ToString() });
            }
            catch (Exception ex)
            {
                MessageBox.Show("getFeatureDictionary Exception: " + ex.Message);
                return;
            }

            if (featureDictionary == null || featureDictionary.Count == 0)
            {
                //MessageBox.Show("No Features in Assembly to Show");
                this.dataGridView2.DataSource = null;
                return;
            }

            if (partEnabledDictionary == null || partEnabledDictionary.Count == 0)
            {
                MessageBox.Show("Data Missing to Show the UI: partEnabledDictionary");
                return;
            }

            DataTable ds = null;
            try
            {
                if (grouper_ManageFeatures != null)
                {
                    ds = (DataTable)grouper_ManageFeatures.GroupingSource.DataSource;
                    //if (ds == null)
                    //    File.AppendAllLines("C:\\Users\\3de.AE-LTC\\3DE_LTC\\log.txt", new string[] { "data source is null" });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("grouper_ManageFeatures Source Exception: " + ex.Message);
                return;
            }
            // logFilePath is Empty -- No Need to LOG
            List<FeatureLine> UIFsList = null;
            if (ds != null)
            {
                try
                {
                    UIFsList = ConvertDataTableToList.ConvertDataTableToFeaturesList(ds, "");
                    //if (UIFsList == null) 
                    //    File.AppendAllLines("C:\\Users\\3de.AE-LTC\\3DE_LTC\\log.txt", new string[] { "UIFsList is null" });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ConvertDataTableToFeaturesList: " + ex.Message);
                    return;
                }
            }
            Dictionary<String, List<FeatureLine>> UIFeatureDictionary = null;
            if (UIFsList != null && UIFsList.Count != 0)
            {
                try
                {
                    UIFeatureDictionary = Utlity.BuildFeatureDictionary(UIFsList, "");
                    //if (UIFeatureDictionary == null)
                    //    File.AppendAllLines("C:\\Users\\3de.AE-LTC\\3DE_LTC\\log.txt", new string[] { "UIFeatureDictionary is null" });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("BuildFeatureDictionary: " + ex.Message);
                    return;
                }
            }
            else
            {
                UIFeatureDictionary = null;
            }

            int childCnt = partEnabledDictionary.Count;
            foreach (String s in partEnabledDictionary.Keys)
            {
                String occName = s;
                bool value = false;
                partEnabledDictionary.TryGetValue(s, out value);
                //File.AppendAllLines("C:\\Users\\3de.AE-LTC\\3DE_LTC\\log.txt", new string[] { "partEnabledDictionary value : " + value.ToString() });
                if (value == true)
                {
                    List<FeatureLine> featuresList = null;

                    if (UIFeatureDictionary != null && UIFeatureDictionary.Count != 0)
                    {
                        if (UIFeatureDictionary.ContainsKey(s) == true)
                        {
                            bool Success = UIFeatureDictionary.TryGetValue(occName, out featuresList);
                            if (featuresList != null && featuresList.Count != 0)
                            {
                                //MessageBox.Show(occName + ":::" + featuresList.Count.ToString());
                            }
                            if (Success == true)
                            {
                                if (featuresList != null && featuresList.Count != 0)
                                {
                                    allFeatureList.AddRange(featuresList);
                                }
                            }
                        }
                        else
                        {
                            bool Success = featureDictionary.TryGetValue(occName, out featuresList);
                            if (featuresList != null && featuresList.Count != 0)
                            {
                                //MessageBox.Show(occName + ":::" + featuresList.Count.ToString());
                            }
                            if (Success == true)
                            {
                                if (featuresList != null && featuresList.Count != 0)
                                {
                                    allFeatureList.AddRange(featuresList);
                                }
                            }
                        }

                    }
                    else
                    {
                        bool Success = featureDictionary.TryGetValue(occName, out featuresList);
                        //if (featuresList != null && featuresList.Count != 0)
                        //{
                        //    MessageBox.Show(occName + ":::" + featuresList.Count.ToString());
                        //}
                        if (Success == true)
                        {
                            if (featuresList != null && featuresList.Count != 0)
                            {
                                allFeatureList.AddRange(featuresList);
                            }
                        }
                    }
                }
            }

            //allFeatureList = SolidEdgeReadFeature.getFeatureLines();
            if (allFeatureList == null || allFeatureList.Count == 0)
            {
                 this.dataGridView2.DataSource = null;
                 //File.AppendAllLines("C:\\Users\\3de.AE-LTC\\3DE_LTC\\log.txt", new string[] { "No Features in Assembly to Show" });
                //MessageBox.Show("No Features in Assembly to Show");
                return;
            }

            List<FeatureLine> newAllFsList = PopulateSKYname(allFeatureList);
            if (newAllFsList == null || newAllFsList.Count == 0)
            {
                MessageBox.Show("No Features to show..");
                return;

            }

            DataTable table = null;
            try
            {
                table = Utlity.ConvertToDataTable(newAllFsList);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ConvertToDataTable: " + ex.Message);
                return;
            }

            if (table == null)
            {
                MessageBox.Show("Unable to Convert Variables to Table");
                return;
            }

            //MessageBox.Show(table.Rows.Count.ToString());            
            this.dataGridView2.DataSource = table;


            try
            {
                applyGroup_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Apply Group to  Feature Table");
                return;

            }


            try
            {
                ApplyAdditionalSettings_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Apply Settings to  Feature Table");
                return;

            }

            try
            {
                hideColumns_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to hide Columns to  Feature Table");
                return;

            }

            try
            {
                AppyHeaderNames_ManageFeatures();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in Setting Column Names in DGV : " + ex.Message);
                return;
            }

            //clearing temp stores
            if (UIFsList != null)
                UIFsList.Clear();
            if (UIFeatureDictionary != null)
                UIFeatureDictionary.Clear();
            if (allFeatureList != null)
                allFeatureList.Clear();

        }

        private void populateDataFromExcel()
        {
            allVarList.Clear();
            //allVarList = SolidEdgeData1.getVariableDetails();
            Dictionary<String, List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
            int childCnt = partEnabledDictionary.Count;
            foreach (String s in partEnabledDictionary.Keys)
            {
                String occName = s;
                bool value = false;
                partEnabledDictionary.TryGetValue(s, out value);
                if (value == true)
                {
                    List<Variable> variablesList = null;
                    bool Success = variableDictionary.TryGetValue(occName, out variablesList);
                    //MessageBox.Show(occName + ":::" + variablesList.Count.ToString());
                    if (Success == true)
                    {
                        allVarList.AddRange(variablesList);
                    }
                }
            }
            List<Variable> newAllVarsList = PopulateSKYname(allVarList);
            if (newAllVarsList == null || newAllVarsList.Count == 0)
            {
                MessageBox.Show("No variables to show..");
                return;

            }

            DataTable table = null;
            table = Utlity.ConvertToDataTable(newAllVarsList);

            //MessageBox.Show(table.Rows.Count.ToString());
            this.dataGridView1.DataSource = table;
            //DgvFilterManager filterManager = new DgvFilterManager(dataGridView1);
            //AppyHeaderNames()

            hideColumns();
            ApplyAdditionalSettings();
            setColumnsSettingsInDGV();
            AppyHeaderNames();
        }


        private void AppyHeaderNames()
        {
            dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.Columns["PartName"].HeaderText = "Part Name";
            this.dataGridView1.Columns["name"].HeaderText = "Name";
            this.dataGridView1.Columns["value"].HeaderText = "Value";
            this.dataGridView1.Columns["rangeLow"].HeaderText = "Range Low";
            this.dataGridView1.Columns["rangeHigh"].HeaderText = "Range High";
            this.dataGridView1.Columns["LOV"].HeaderText = "LOV";
            this.dataGridView1.Columns["systemName"].HeaderText = "System Name";
            this.dataGridView1.Columns["unit"].HeaderText = "Unit";
            this.dataGridView1.Columns["rangeCondition"].HeaderText = "Range Condition";
            this.dataGridView1.Columns["variableType"].HeaderText = "Variable Type";
            this.dataGridView1.Columns["AddVarToTemplate"].HeaderText = "Add Variable(Y/N)";
            this.dataGridView1.Columns["AddPartToTemplate"].HeaderText = "Add Part(Y/N)";

        }

        private void AppyHeaderNames_ManageFeatures()
        {
            dataGridView2.AutoGenerateColumns = false;
            this.dataGridView2.Columns["FeatureName"].HeaderText = "Internal Name";
            this.dataGridView2.Columns["EdgeBarName"].HeaderText = "Display Name";

            this.dataGridView2.Columns["SuppressionEnabled"].HeaderText = "Suppression Enabled";

        }
        private void setColumnsSettingsInDGV()
        {
            //set autosize mode
            //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //datagrid has calculated it's widths so we can store them
            for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                //store autosized widths
                int colw = dataGridView1.Columns[i].Width;
                //remove autosizing
                dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                //set width to calculated by autosize
                dataGridView1.Columns[i].Width = colw;
            }
        }


        //optionally, you can customize the grouping display by subscribing to the DisplayGroup event
        void grouper_DisplayGroup(object sender, Subro.Controls.GroupDisplayEventArgs e)
        {
            e.BackColor = (e.Group.GroupIndex % 2) == 0 ? Color.LightBlue : Color.White;
            //e.BackColor = Color.LightGray;
            //e.Header = "[" + e.Header + "], grp: " + e.Group.GroupIndex;
            e.Header = "";
            e.DisplayValue = e.DisplayValue;
            e.Summary = "(" + e.Group.Count + " )";
            e.ForeColor = Color.Black;
        }

        private void TabControlSettings()
        {
            tabControl1.AutoSize = true;
            tabControl1.TabPages[0].AutoSize = true;
            tabControl1.TabPages[1].AutoSize = true;
        }

        private void ApplyAdditionalSettings()
        {
            this.dataGridView1.Columns["PartName"].ReadOnly = true;
            this.dataGridView1.Columns["name"].ReadOnly = false;
            this.dataGridView1.Columns["systemName"].ReadOnly = true;
            this.dataGridView1.Columns["unit"].ReadOnly = true;
            this.dataGridView1.Columns["rangeCondition"].ReadOnly = true;
            this.dataGridView1.Columns["Formula"].ReadOnly = false;
            this.dataGridView1.Columns["SKYname"].ReadOnly = false;

            
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Azure;
            //dataGridView1.DefaultCellStyle.SelectionBackColor = Color.White;
            //this.dataGridView1.Columns["PartName"].Visible = false;
            //this.dataGridView1.Columns["AddPartToTemplate"].Visible = false;
        }

        private void hideColumns()
        {
            this.dataGridView1.Columns["name"].Frozen = true;

            
            this.dataGridView1.Columns["DefaultValue"].Visible = false;
            this.dataGridView1.Columns["systemName"].Visible = false;
            this.dataGridView1.Columns["unit"].Visible = true;
            this.dataGridView1.Columns["rangeCondition"].Visible = false;
            this.dataGridView1.Columns["UnitType"].Visible = false;
            //this.dataGridView1.Columns["PartName"].Visible = false;
            this.dataGridView1.Columns["variableType"].Visible = true;
            this.dataGridView1.Columns["AddVarToTemplate"].Visible = true;
            this.dataGridView1.Columns["AddPartToTemplate"].Visible = false;
            this.dataGridView1.Columns["SKYname"].Visible = false;
        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(tabControl1.SelectedTab.Text);
            //if (tabControl1.SelectedTab.Text == "Manage Components")
            //{                              

            //}
            if (tabControl1.SelectedTab.Text == "Manage Variables")
            {

                getSelectedComponentsFromTreeView();
                //populateDataFromExcel();
                ClubUIAndDataFromSE();
                //applyGroup();
                grouper.RemoveGrouping();
                //dataGridView1.Columns["PartName"].Visible = false;

            }
            if (tabControl1.SelectedTab.Text == "Manage Features")
            {
                getSelectedComponentsFromTreeView();
                //PopulateFeatureDataFromExcel();
                try
                {
                    ClubUIAndFeatureDataFromSE();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //grouper_ManageFeatures.RemoveGrouping();
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(toolStripMenuItem1.Text);
            if (m_OldSelectNode != null)
            {
                m_OldSelectNode.Checked = true;
                MyCustomDialog4View.TraverseNodes(m_OldSelectNode, true);
                //int childNodesCnt = treeView1.Nodes[0].Nodes.Count;

                //for (int i = 0; i < childNodesCnt; i++)
                //{
                //    treeView1.Nodes[0].Nodes[i].Checked = true;
                //}
            }

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

            //MessageBox.Show(toolStripMenuItem2.Text);

            //treeView1.Nodes[0].Checked = false;
            //int childNodesCnt = treeView1.Nodes[0].Nodes.Count;

            //for (int i = 0; i < childNodesCnt; i++)
            //{
            //    treeView1.Nodes[0].Nodes[i].Checked = false;
            //}

            if (m_OldSelectNode != null)
            {
                m_OldSelectNode.Checked = false;
                MyCustomDialog3View.TraverseNodes(m_OldSelectNode, false);
            }

        }

        private TreeNode m_OldSelectNode;
        private void treeView1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            // Show menu only if the right mouse button is clicked.
            if (e.Button == MouseButtons.Right)
            {

                // Point where the mouse is clicked.
                Point p = new Point(e.X, e.Y);

                // Get the node that the user has clicked.
                TreeNode node = treeView1.GetNodeAt(p);
                if (node != null)
                {

                    // Select the node the user has clicked.
                    // The node appears selected until the menu is displayed on the screen.
                    m_OldSelectNode = treeView1.SelectedNode;
                    treeView1.SelectedNode = node;

                    //MessageBox.Show(treeView1.Nodes[0].Text);
                    //MessageBox.Show(node.Text);
                    // Find the appropriate ContextMenu depending on the selected node.
                    if (treeView1.SelectedNode.Nodes != null && treeView1.SelectedNode.Nodes.Count != 0)
                    {
                        contextMenuStrip1.Show(treeView1, p);
                    }


                    // Highlight the selected node.
                    treeView1.SelectedNode = m_OldSelectNode;
                    //m_OldSelectNode = null;
                }
            }
        }

        

        
        // Save
        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Save Clicked");

            
            // CTE uses SolidEdgeData1 initially to SET assembly FILE and OPEN THE UI
            String assemFileName = SolidEdgeData1.getAssemblyFileName();

            if (assemFileName == null || assemFileName.Equals("") == true)
            {
                MessageBox.Show("Could not Find Root Assembly File");
                return;
            }
            String AssemblyStageDir = System.IO.Path.GetDirectoryName(assemFileName);
            String LogStageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(assemFileName) + "_" + "CTE" + ".txt");
            String outputXLfileName = System.IO.Path.Combine(AssemblyStageDir, System.IO.Path.GetFileNameWithoutExtension(assemFileName) + ".xlsx");
            Console.WriteLine("Opening LogFile @ {0} " + logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            //1 - Convert the Updated DGV to Table/List.    
            DataTable ds = null;
            if (grouper != null)
            {
                ds = (DataTable)grouper.GroupingSource.DataSource;
            }
            else
            {
                MessageBox.Show("No Data in Variable Manager to Save");
                //this.progressBar2.Visible = false;
                return;
            }
            Utlity.Log(ds.Rows.Count.ToString(), logFilePath);
            List<Variable> variablesList = new List<Variable>();
            foreach (DataRow row in ds.Rows)
            {

                //if (row.IsNewRow == true) continue;
                //if (row.DataBoundItem is Subro.Controls.GroupRow)
                //{
                //    continue;
                //}
                //Utlity.Log("Reading : " + (String)row["PartName"].ToString(), logFilePath);
                Variable varr = new Variable();
                varr.PartName = Utlity.getValue(row, "PartName", logFilePath);
                Utlity.Log("varr.PartName " + varr.PartName, logFilePath);
                varr.name = Utlity.getValue(row, "name", logFilePath);
                Utlity.Log("varr.name " + varr.name, logFilePath);
                varr.systemName = Utlity.getValue(row, "systemName", logFilePath);
                varr.value = Utlity.getValue(row, "value", logFilePath);
                Utlity.Log("varr.value " + varr.value, logFilePath);
                varr.DefaultValue = varr.value;
                Utlity.Log("varr.DefaultValue " + varr.DefaultValue, logFilePath);
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
                varr.AddPartToTemplate = "Y";
                varr.variableType = Utlity.getValue(row, "variableType", logFilePath);
                varr.UnitType = Utlity.getValue(row, "UnitType", logFilePath);
                variablesList.Add(varr);
            }

            Utlity.Log("Adding Parts which are Hidden", logFilePath);
            // **ADD all occurences which are HIDDEN but Still Needed in the XL Sheet ***** START//
            foreach (String s in partEnabledDictionary.Keys)
            {
                bool value = false;
                partEnabledDictionary.TryGetValue(s, out value);
                if (value == false)
                {
                    Dictionary<String, List<Variable>> variableDictionary = ExcelData.getVariableDictionary();
                    if (variableDictionary.Count != 0)
                    {
                        List<Variable> PartHiddenvariableList = null;
                        bool Success = variableDictionary.TryGetValue(s, out PartHiddenvariableList);

                        if (Success == true)
                        {
                            if (PartHiddenvariableList != null && variableDictionary.Count != 0)
                            {
                                //Utlity.Log("Hidden Part: " + s, logFilePath);
                                foreach (Variable varr in PartHiddenvariableList)
                                {
                                    varr.AddPartToTemplate = "N";
                                }
                                variablesList.AddRange(PartHiddenvariableList);
                            }

                        }
                    }

                }
            }

            Utlity.Log("Collecting Modified Features ", logFilePath);
            DataTable ds1 = null;
            List<FeatureLine> featureLineList = new List<FeatureLine>();
            if (grouper_ManageFeatures != null)
            {
                try
                {
                    //ds1 = ((DataView)grouper_ManageFeatures.GroupingSource.DataSource).Table;
                    ds1 = (DataTable)grouper_ManageFeatures.GroupingSource.DataSource;
                }
                catch (Exception ex)
                {
                    Utlity.Log("grouper_ManageFeatures Data Source Exception: " + ex.Message, logFilePath);
                    return;
                }
                //if (ds1 == null || ds1.Rows.Count == 0)
                //{
                //    Utlity.Log("No Features in the Source ", logFilePath);
                //    MessageBox.Show("Select the Features Tab Before Saving the Template");
                //    return;
                //}
                try
                {
                    if (ds1 != null || ds1.Rows.Count != 0)
                    {
                        featureLineList = ConvertDataTableToList.ConvertDataTableToFeaturesList(ds1, logFilePath);
                    }
                    else
                    {
                        Utlity.Log("No Features in the Source ", logFilePath);
                    }
                }
                catch (Exception ex)
                {
                    Utlity.Log("No Features in the Source " + ex.Message, logFilePath);
                }
                //if (featureLineList == null || featureLineList.Count == 0)
                //{
                //    return;
                //}
            }
            else
            {
                Utlity.Log("No Features in the Source ", logFilePath);
            }

            Utlity.Log("Add All Features that were Not Selected By User ", logFilePath);
            // **ADD Feature Info of occurences which are HIDDEN but Still Needed in the XL Sheet ***** START//
            foreach (String s in partEnabledDictionary.Keys)
            {
                Dictionary<String, List<FeatureLine>> featureDictionary = ExcelReadFeatures.getFeatureDictionary();
                bool value = false;
                partEnabledDictionary.TryGetValue(s, out value);
                if (value == false)
                {
                    // Change To ExcelReadFeature
                    
                    if (featureDictionary != null && featureDictionary.Count != 0)
                    {
                        List<FeatureLine> PartHiddenFsList = null;
                        bool Success = featureDictionary.TryGetValue(s, out PartHiddenFsList);

                        if (Success == true)
                        {
                            if (PartHiddenFsList != null && PartHiddenFsList.Count != 0)
                            {
                                //Utlity.Log("Hidden Part: " + s, logFilePath);
                                featureLineList.AddRange(PartHiddenFsList);
                            }

                        }
                    }

                }
                else
                {
                    if (ds1 == null)
                    {                        
                        if (featureDictionary != null && featureDictionary.Count != 0)
                        {
                            List<FeatureLine> PartHiddenFsList = null;
                            bool Success = featureDictionary.TryGetValue(s, out PartHiddenFsList);

                            if (Success == true)
                            {
                                if (PartHiddenFsList != null && PartHiddenFsList.Count != 0)
                                {
                                    //Utlity.Log("Hidden Part: " + s, logFilePath);
                                    featureLineList.AddRange(PartHiddenFsList);
                                }

                            }
                        }
                    }
                }
            }
            // **ADD Feature Info of occurences which are HIDDEN but Still Needed in the XL Sheet ***** END//
            
            if (System.IO.File.Exists(outputXLfileName) == false)
            {
                MessageBox.Show("CTE - Excel File does not Exist");
                return;
            }
            if (variablesList.Count == 0)
            {
                MessageBox.Show("No Variables To Save to Excel");
                return;

            }

            //this.Dispose();
            List<object> arguments = new List<object>();
            arguments.Add(outputXLfileName);
            arguments.Add(variablesList);
            arguments.Add(logFilePath);
            arguments.Add(partEnabledDictionary);
            arguments.Add(featureLineList);
            if (backgroundWorker1.IsBusy != true)
            {
                this.progressBar2.Visible = true;
                MyCustomDialog4View.EnableDisableAllControls(false, this);
                //Utlity.Log("Calling Background Worker", logFilePath);
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker is Busy, Save in CTE", logFilePath);
            }
        }

        // v1.0.0.5, 18/11/2018 - change to Updating Only the Delta Portions.
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //MessageBox.Show("Inside BackGround Worker");
            List<object> genericlist = e.Argument as List<object>;
            String OldoutputXLfileName = (String)genericlist[0];
            //MessageBox.Show(outputXLfileName);            
            List<Variable> variablesList = (List<Variable>)genericlist[1];
            //MessageBox.Show(variablesList.Count.ToString());
            String logFilePath = (String)genericlist[2];
            //MessageBox.Show(logFilePath);
            Dictionary<String, bool> partEnabledDictionary = (Dictionary<String, bool>)genericlist[3];
            //MessageBox.Show(partEnabledDictionary.Count.ToString());
            List<FeatureLine> featureLineList = (List<FeatureLine>)genericlist[4];
            //MessageBox.Show(featureLineList.Count.ToString());

            Utlity.Log("Output XL File: " + OldoutputXLfileName, logFilePath);
            Utlity.Log("Variable Count: " + variablesList.Count, logFilePath);
            String fileNameWithoutExtn = System.IO.Path.GetFileNameWithoutExtension(OldoutputXLfileName);
            String DirectoryName = System.IO.Path.GetDirectoryName(OldoutputXLfileName);



            String NewTemplateXLFilePath = System.IO.Path.Combine(DirectoryName, fileNameWithoutExtn + "_New" + ".xlsx");

            if (System.IO.File.Exists(NewTemplateXLFilePath) == true)
            {
                // If the New File Already Exists.
                Utlity.Log("Delete New XL File: " + NewTemplateXLFilePath, logFilePath);
                try
                {
                    System.IO.File.Delete(NewTemplateXLFilePath);
                }
                catch (Exception ex)
                {
                    Utlity.Log("Could not Delete XL File:  " + NewTemplateXLFilePath + " " + ex.Message, logFilePath);
                }
            }

            string tc_mode = Utlity.getManageMode(logFilePath);
            string assemblyFileName = Path.ChangeExtension(OldoutputXLfileName, ".asm");
            if (tc_mode.Equals("YES", StringComparison.OrdinalIgnoreCase))
            {
                SEECAdaptor.LoginToTeamcenter(logFilePath);
                string revID = SEECAdaptor.getRevisionID(assemblyFileName);
                SEECAdaptor.DownloadFilesIntoCache(Path.GetFileNameWithoutExtension(assemblyFileName), revID, logFilePath);
                Utlity.Log(string.Concat("open solid edge draft and register the scale factor ", assemblyFileName), logFilePath, null);
                SEECAdaptor.collectDraftFilesFromCacheAndRegisterScaleFactor(logFilePath);
            }

            ExcelInterface.SaveToXL(NewTemplateXLFilePath, variablesList, partEnabledDictionary, logFilePath, "CTE", featureLineList);
            CTEExcelDeltaUpdate_1.ProcessNewTemplate(NewTemplateXLFilePath, OldoutputXLfileName, logFilePath);

            if (System.IO.File.Exists(OldoutputXLfileName) == true)
            {
                // If the Old File Already Exists.
                Utlity.Log("Delete Old XL File: " + OldoutputXLfileName, logFilePath);
                try
                {
                    System.IO.File.Delete(OldoutputXLfileName);
                }
                catch (Exception ex)
                {
                    Utlity.Log("Could not Delete XL File:  " + OldoutputXLfileName + " " + ex.Message, logFilePath);
                }
            }

            // Rename the New File to the Old File Now
            if (System.IO.File.Exists(NewTemplateXLFilePath) == true)
            {
                try
                {

                    System.IO.File.Move(NewTemplateXLFilePath, OldoutputXLfileName);
                }
                catch (Exception ex)
                {
                    Utlity.Log("Could not Move XL File:  " + NewTemplateXLFilePath + " " + ex.Message, logFilePath);
                }
            }


            genericlist.Clear();
            List<object> arguments = new List<object>();
            arguments.Add(OldoutputXLfileName);
            arguments.Add(variablesList);
            arguments.Add(logFilePath);
            arguments.Add(partEnabledDictionary);
            arguments.Add(featureLineList);
            ExcelUtils.KillExcelFileProcess(logFilePath);
            tc_mode = Utlity.getManageMode(logFilePath);

            // Murali - 25-NOV-2024 - SOA Decustomization - Start
            //if (tc_mode.Equals("YES", StringComparison.OrdinalIgnoreCase) == true)
            //{
            //    TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
            //    TcAdaptor.TcAdaptor_Init(logFilePath);
            //    Utlity.Log("upload template excel to teamcenter.." + OldoutputXLfileName, logFilePath);
            //    TcAdaptor.uploadExcelToTC(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, OldoutputXLfileName, logFilePath);
            //    Utlity.Log("check in all SE documents in cache to teamcenter.." + OldoutputXLfileName, logFilePath);
            //    SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath);
            //    //TcAdaptor.logout(logFilePath);

            //    assemblyFileName = Path.ChangeExtension(OldoutputXLfileName, ".asm");

            //    //SanitizeXL_PostUpload_Logic.traverseAssembly(assemblyFileName, logFilePath);
            //    if (partEnabledDictionary.Count > 0)
            //    {
            //        Utlity.Log("update the variable part type in Teamcenter for.." + assemblyFileName, logFilePath);
            //        //TcAdaptor.login(loginFromSE.userName, loginFromSE.password, loginFromSE.group, loginFromSE.role, logFilePath);
            //        //TcAdaptor.TcAdaptor_Init(logFilePath);
            //        SanitizeXL_PostUpload_Logic.updateSuffixProperty(partEnabledDictionary, logFilePath);

            //    }
            //    TcAdaptor.logout(logFilePath);
            //}
            // Murali - 25-NOV-2024 - SOA Decustomization - End


            if (tc_mode.Equals("YES", StringComparison.OrdinalIgnoreCase) == true)
            {
                SEECAdaptor.LoginToTeamcenter(logFilePath);

                string bStrCurrentUser = null;
                SEECAdaptor.getSEECObject().GetCurrentUserName(out bStrCurrentUser);

                String password = bStrCurrentUser;


                Utlity.Log("Logging into TC..for TVS..NOV 2024", logFilePath);
                Utlity.Log("Attempting SOA Login using following credentials: ", logFilePath);
                Utlity.Log("ID=" + bStrCurrentUser, logFilePath);
                Utlity.Log("Group=DBA", logFilePath);
                Utlity.Log("Role=dba", logFilePath);
                TcAdaptor.login(bStrCurrentUser, password, "DBA", "dba", logFilePath);
                Utlity.Log("Initializing TC Services..", logFilePath);
                TcAdaptor.TcAdaptor_Init(logFilePath);
                TcAdaptor.uploadExcelToTC(bStrCurrentUser, password, "DBA", "dba", OldoutputXLfileName, logFilePath);

                Utlity.Log("check in all SE documents in cache to teamcenter.." + OldoutputXLfileName, logFilePath);
                SEECAdaptor.CheckInSEDocumentsToTeamcenter(logFilePath);
                assemblyFileName = Path.ChangeExtension(OldoutputXLfileName, ".asm");

                //SanitizeXL_PostUpload_Logic.traverseAssembly(assemblyFileName, logFilePath);
                if (partEnabledDictionary.Count > 0)
                {
                    Utlity.Log("update the variable part type in Teamcenter for.." + assemblyFileName, logFilePath);
                    SanitizeXL_PostUpload_Logic.updateSuffixProperty(partEnabledDictionary, logFilePath);
                }

                TcAdaptor.logout(logFilePath);
            }


            if (tc_mode.Equals("YES", StringComparison.OrdinalIgnoreCase) == true)
            {
                MessageBox.Show("Saved the Template Successfully to Teamcenter: " + OldoutputXLfileName);
            }
            else
            {
                MessageBox.Show("Saved the Template Successfully: " + OldoutputXLfileName);
            }

            //MessageBox.Show(arguments.Count.ToString());
            e.Result = "OK";
            return;
        }


        // v1.0.0.4, Commented to change the Flow on 18/11/2018
        //private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    //MessageBox.Show("Inside BackGround Worker");
        //    List<object> genericlist = e.Argument as List<object>;
        //    String outputXLfileName = (String)genericlist[0];
        //    //MessageBox.Show(outputXLfileName);            
        //    List<Variable> variablesList = (List<Variable>)genericlist[1];
        //    //MessageBox.Show(variablesList.Count.ToString());
        //    String logFilePath = (String)genericlist[2];
        //    //MessageBox.Show(logFilePath);
        //    Dictionary<String, bool> partEnabledDictionary = (Dictionary<String, bool>)genericlist[3];
        //    //MessageBox.Show(partEnabledDictionary.Count.ToString());
        //    List<FeatureLine> featureLineList = (List<FeatureLine>)genericlist[4];
        //    //MessageBox.Show(featureLineList.Count.ToString());

        //    Utlity.Log("Output XL File: " + outputXLfileName, logFilePath);
        //    Utlity.Log("Variable Count: " + variablesList.Count, logFilePath);

        //    if (System.IO.File.Exists(outputXLfileName) == true)
        //    {
        //        Utlity.Log("Delete XL File: " + outputXLfileName, logFilePath);
        //        System.IO.File.Delete(outputXLfileName);
        //    }
        //    ExcelInterface.SaveToXL(outputXLfileName, variablesList, partEnabledDictionary, logFilePath, "CTE", featureLineList);
        //    genericlist.Clear();
        //    List<object> arguments = new List<object>();
        //    arguments.Add(outputXLfileName);
        //    arguments.Add(variablesList);
        //    arguments.Add(logFilePath);
        //    arguments.Add(partEnabledDictionary);
        //    arguments.Add(featureLineList);
        //    MessageBox.Show("Saved the Template Successfully: " + outputXLfileName);
        //    //MessageBox.Show(arguments.Count.ToString());
        //    e.Result = "OK";
        //    return;
        //}

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Result.ToString().Equals("OK")) {
                //MessageBox.Show(e.Result.ToString());
                //this.Dispose();
                this.DialogResult = DialogResult.OK;
                //SolidEdgeFramework.Application application = SE_SESSION.getSolidEdgeSession();
                //if (application != null)
                //{
                //    this.Dispose();
                //    application.Activate();
                //}
                
                return;
            }
            //if (e.Error != null)
            //{
            //    MessageBox.Show(e.Error.Message);
            //    return;
            //}
            //MessageBox.Show("backgroundWorker1_RunWorkerCompleted");            
            //List<object> genericlist = (List<object>)e.Result;
            //if (genericlist == null || genericlist.Count == 0)
            //{
            //    return;
            //}
            //String outputXLfileName = (String)genericlist[0];
            //List<Variable> variablesList = (List<Variable>)genericlist[1];
            //String logFilePath = (String)genericlist[2];
            //Dictionary<String, bool> partEnabledDictionary = (Dictionary<String, bool>)genericlist[3];


            //MessageBox.Show("Saved the Template Successfully: " + outputXLfileName);
            //Utlity.Log("VariablesList Count:" + variablesList.Count, logFilePath);
            //Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            //Utlity.Log("-----------------------------------------------------------------", logFilePath);
            //this.Close();
        }

        private TreeNode previousNode = null;

        private void treeView1_NodeMouseHover(object sender, TreeNodeMouseHoverEventArgs e)
        {
            
            if ( previousNode != null) {
                previousNode.ForeColor = Color.Empty;
                previousNode.BackColor = Color.Empty;                
            }

            e.Node.ForeColor = Color.FromKnownColor(KnownColor.HighlightText);
            e.Node.BackColor = Color.FromKnownColor(KnownColor.Highlight) ;  
            previousNode = e.Node ;
        }

      

        //private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        //{
        //    if (e.Node != null)
        //    {
        //        String assemblyFileName = treeView1.Nodes[0].Text;
        //        SolidEdgeHighLighter.HighlightOccurence(assemblyFileName, e.Node.Text);
        //    }

        //}

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Action.Equals(TreeViewAction.Unknown))
            {
                return;
            }

            if (e.Node != null)
            {
                String assemblyFileName = treeView1.Nodes[0].Text;
                SolidEdgeHighLighter.HighlightOccurence(assemblyFileName, e.Node.Text);
            }

        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Action.Equals(TreeViewAction.Unknown))
            {
                return;
            }
            String nodeText = e.Node.Text;
            bool checkStatus = e.Node.Checked;


            String LogStageDir = Utlity.CreateLogDirectory();
            //String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(nodeText) + "_" + "TREEVIEW" + ".txt");
            String logFilePath = "";
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            if (e.Node != null)
            {
                if (nodeText != null && nodeText.Equals("") == false)
                {
                    try
                    {
                        Utlity.Log("nodeText " + nodeText, logFilePath);
                        Utlity.Log("checkStatus " + checkStatus, logFilePath);
                        MyCustomDialog4View.selectAllNodesWithText(treeView1, nodeText, checkStatus, logFilePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
        }

        private void FeatureValidate_Click(object sender, EventArgs e)
        {
            String assemFileName = SolidEdgeData1.getAssemblyFileName();
            if (assemFileName == null || assemFileName.Equals("") == true)
            {
                return;
            }
            String AssemblyStageDir = System.IO.Path.GetDirectoryName(assemFileName);
            String LogStageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(assemFileName) + "_" + "FeatureFormulaValidate" + ".txt");
            Console.WriteLine("Opening LogFile @ {0} " + logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            DataTable ds1 = null;
            try
            {
                //ds1 = ((DataView)grouper_ManageFeatures.GroupingSource.DataSource).Table;
                ds1 = (DataTable)grouper_ManageFeatures.GroupingSource.DataSource;
            }
            catch (Exception ex)
            {
                Utlity.Log("grouper_ManageFeatures Data Source Exception: " + ex.Message, logFilePath);
                return;
            }
            if (ds1 == null || ds1.Rows.Count == 0)
            {
                Utlity.Log("No Features in the Source ", logFilePath);
                return;
            }
            List<FeatureLine> featureLineList = ConvertDataTableToList.ConvertDataTableToFeaturesList(ds1, logFilePath);

            // print the Feature Lines from the UI
            Utlity.Log("featureLineList: " + featureLineList.Count, logFilePath);

            // Added on 26-12, Any New Features that are Added After TVS is not Reflecting in the CTE UI.
            // Read the Features from Solid Edge and Those Features that Are Missing in featreLineList can be Added into the List.
            List<FeatureLine> featureLineListFromSolidEdge = readLatestSolidEdgeFeatures(logFilePath);

            // print the Feature Lines from the Solid Edge
            Utlity.Log("featureLineListFromSolidEdge: " + featureLineListFromSolidEdge.Count, logFilePath);

            // 01-01-2019 , Add Features into The List, These New Features are Added to Solid Edge After TVS Excel is Created.
            featureLineList = compareAndAddRemoveFeaturelinesInList(featureLineList, featureLineListFromSolidEdge, logFilePath);

            DataTable ds = null;
            try
            {
                ds = (DataTable)grouper.GroupingSource.DataSource;
                //ds = ((DataView)grouper.GroupingSource.DataSource).Table;
            }
            catch (Exception ex)
            {
                Utlity.Log("grouper Data Source Exception: " + ex.Message, logFilePath);
                return;
            }


            if (ds == null || ds.Rows.Count == 0)
            {
                Utlity.Log("No Variables in the Source ", logFilePath);
                return;
            }
            List<Variable> variablesList = ConvertDataTableToList.ConvertDataTableToVariablesList(ds, logFilePath);


            List<object> arguments = new List<object>();
            arguments.Add(assemFileName);
            arguments.Add(featureLineList);
            arguments.Add(variablesList);
            arguments.Add(logFilePath);
            arguments.Add(partEnabledDictionary);
            if (backgroundWorker2.IsBusy != true)
            {
                MyCustomDialog4View.EnableDisableAllControls(false, this);
                backgroundWorker2.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker 2 is Busy, Validate Feature Formula Sync", logFilePath);
            }
        }

        // 1 Jan 2019 - New FeatureLine from SolidEdge is Added into the FeatureLineList.
        // As a Result the new FeatureLine is Displayed in the UI and would be Taken into Excel Sheet, During CTE
        private List<FeatureLine> compareAndAddRemoveFeaturelinesInList(List<FeatureLine> featureLineList, List<FeatureLine> featureLineListFromSolidEdge, String logFilePath)
        {            

            foreach (FeatureLine fl in featureLineListFromSolidEdge)
            {
                var result = featureLineList.Find(fl1 => (fl1.PartName.Equals(fl.PartName, StringComparison.OrdinalIgnoreCase)) &&
                    (fl1.FeatureName.Equals(fl.FeatureName, StringComparison.OrdinalIgnoreCase)));
                if (result == null)
                {
                    Utlity.Log("compareAndAddRemoveFeaturelinesInList: Added new FeatureLine, " + fl.PartName + " " + fl.FeatureName, logFilePath);
                    featureLineList.Add(fl);
                }
                   
            }


            for (int x = featureLineList.Count - 1; x > -1; x-- )
            {
                FeatureLine fl = featureLineList[x];
                var result = featureLineListFromSolidEdge.Find(fl1 => (fl1.PartName.Equals(fl.PartName, StringComparison.OrdinalIgnoreCase)) &&
                    (fl1.FeatureName.Equals(fl.FeatureName, StringComparison.OrdinalIgnoreCase)));

                if (result == null)
                {
                    Utlity.Log("compareAndAddRemoveFeaturelinesInList: Removing FeatureLine, " + fl.PartName + " " + fl.FeatureName, logFilePath);
                    try
                    {
                        featureLineList.RemoveAt(x);
                    }
                    catch (Exception ex)
                    {
                        Utlity.Log("compareAndAddRemoveFeaturelinesInList: Removing FeatureLine Ex: , " + fl.PartName + " " + fl.FeatureName + " " + ex.Message, logFilePath);
                        continue;
                    }
                }
            }

            //Utlity.Log("compareAndAddRemoveFeaturelinesInList: Sorting the Final FeatureLine List..", logFilePath);
            featureLineList = featureLineList.OrderBy(p => p.PartName).ToList();
            return featureLineList;
        }

        //  Added on 26-Dec, CTE is Actually not Reflecting New Features if they are Added after TVS
        private List<FeatureLine> readLatestSolidEdgeFeatures(string logFilePath)
        {
            List<FeatureLine> featureLineListFromSolidEdge = new List<FeatureLine>();
            // Start a Thread here, Since SolidEdge Functionality Runs Only in STA MODE.
            try
            {
                Utlity.Log("readLatestSolidEdgeFeatures @ " + System.DateTime.Now.ToString(), logFilePath);                
                Thread myThread = new Thread(() => SolidEdgeReadFeature.readFeatures(logFilePath, "CTE")); 
                myThread.SetApartmentState(ApartmentState.STA);
                myThread.Start();
                myThread.Join();
            }
            catch (Exception ex)
            {
                this.DialogResult = DialogResult.Cancel;             
                Utlity.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);
            }

            featureLineListFromSolidEdge = SolidEdgeReadFeature.getFeatureLines();
            return featureLineListFromSolidEdge;
            
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;

            String assemblyFileName = (String)genericlist[0];
            List<FeatureLine> featureLinesList = (List<FeatureLine>)genericlist[1];
            List<Variable> variablesList = (List<Variable>)genericlist[2];
            String logFilePath = (String)genericlist[3];
            Dictionary<String, bool> partEnabledDict = (Dictionary<String, bool>)genericlist[4];

            Utlity.Log("validateFeatureFormula: ", logFilePath);
            backgroundWorker2.ReportProgress(10);
            List<FeatureLine> updatedFsList = FeatureFormulaParser.validateFeatureFormula(variablesList, featureLinesList, logFilePath);
            backgroundWorker2.ReportProgress(50);
            // Start a Thread here, Since SolidEdge Functionality Runs Only in STA MODE.
            try
            {
                Utlity.Log("SolidEdgeSetFeatureState: setFeatures: " + System.DateTime.Now.ToString(), logFilePath);
                Thread myThread = new Thread(() => SolidEdgeSetFeatureState.setFeatures(logFilePath, updatedFsList, partEnabledDict,"CTE"));
                myThread.SetApartmentState(ApartmentState.STA);
                myThread.Start();
                myThread.Join();
            }
            catch (Exception ex)
            {
                MyCustomDialog4View.EnableDisableAllControls(true, this);
                Utlity.Log("SolidEdgeReadFeature, readFeatures: " + ex.Message, logFilePath);
                e.Result = null;
                return;
            }

            backgroundWorker2.ReportProgress(95);
            if (updatedFsList != null)
            {
                Utlity.Log("updatedFsList" + updatedFsList.Count, logFilePath);
            }
            if (partEnabledDict != null)
            {
                Utlity.Log("partEnabledDict" + partEnabledDict.Count, logFilePath);
            }
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);

            genericlist[0] = assemblyFileName;
            genericlist[1] = updatedFsList;
            genericlist[2] = variablesList;
            genericlist[3] = logFilePath;
            genericlist[4] = partEnabledDict;
            e.Result = genericlist;

        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = (List<object>)e.Result;
            String logFilePath = null;
            if (genericlist == null || genericlist.Count == 0 || genericlist[1] == null)
            {
                MessageBox.Show("Failure in finding features list. Is path correctly mapped in the template excel");
                MyCustomDialog4View.EnableDisableAllControls(true, this);
                if (genericlist[1] == null)
                {
                    logFilePath = (String)genericlist[3];
                    Utlity.Log("UpdateRowsInManageFeatures : " + "FeatureLines to Update in UI is Empty...", logFilePath);
                }
                this.progressBar1.Value = 0;
                this.progressBar1.Visible = false;
                return;
            }
            String assemblyFileName = (String)genericlist[0];
            logFilePath = (String)genericlist[3];
            if (genericlist[1] != null)
            {
                List<FeatureLine> featureLinesList = (List<FeatureLine>)genericlist[1];
                List<Variable> variablesList = (List<Variable>)genericlist[2];               
                Dictionary<String, bool> partEnabledDict = (Dictionary<String, bool>)genericlist[4];
                Utlity.Log("UpdateRowsInManageFeatures : " + featureLinesList.Count, logFilePath);
                MyCustomDialog4View.EnableDisableAllControls(true, this);
                UpdateRowsInManageFeatures(featureLinesList);
            }
            else
            {
                MyCustomDialog4View.EnableDisableAllControls(true, this);
            }
        }

        private void UpdateRowsInManageFeatures(List<FeatureLine> featureLinesList) 
        {
            this.progressBar1.Visible = false;
            if (featureLinesList == null || featureLinesList.Count == 0)
            {
                //MessageBox.Show("No Features in Assembly to Show");
                return;
            }

            List<FeatureLine> newAllFsList = PopulateSKYname(featureLinesList);
            if (newAllFsList == null || newAllFsList.Count == 0)
            {
                MessageBox.Show("No Features to show..");
                return;

            }

            DataTable table = null;
            table = Utlity.ConvertToDataTable(newAllFsList);
            if (table == null)
            {
                MessageBox.Show("Unable to Convert Variables to Table");
                return;
            }

            //MessageBox.Show(table.Rows.Count.ToString());            
            this.dataGridView2.DataSource = table;

        }

        private void FormulaSave_Click(object sender, EventArgs e)
        {
            //this.Dispose();
            String assemFileName = SolidEdgeData1.getAssemblyFileName();
            if (assemFileName == null || assemFileName.Equals("") == true)
            {
                return;
            }
            String AssemblyStageDir = System.IO.Path.GetDirectoryName(assemFileName);
            String LogStageDir = Utlity.CreateLogDirectory();
            String logFilePath = System.IO.Path.Combine(LogStageDir, System.IO.Path.GetFileNameWithoutExtension(assemFileName) + "_" + "SyncFormula" + ".txt");
            String outputXLfileName = System.IO.Path.Combine(AssemblyStageDir, System.IO.Path.GetFileNameWithoutExtension(assemFileName) + ".xlsx");
            Console.WriteLine("Opening LogFile @ {0} " + logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);

            DataTable ds = null;
            try
            {
                ds = (DataTable)grouper.GroupingSource.DataSource;
            }
            catch (Exception ex)
            {
                Utlity.Log("grouper Exception " + ex.Message, logFilePath);
                return;
            }
            //Utlity.Log("ds.Rows.Count.ToString(): " + ds.Rows.Count.ToString(), logFilePath);

            //1 - Convert the Updated DGV to Table/List.            
            List<Variable> variablesList = new List<Variable>();
            //Utlity.Log("Total Rows in Screen: " + ds.Rows.Count.ToString(), logFilePath);
            //Utlity.Log("Modified Rows in Screen: " + ds.GetChanges().Rows.Count.ToString(), logFilePath);
            if (ds.Rows == null)
            {
                MessageBox.Show("Review the Data before Submit");
                return;
            }
            foreach (DataRow row in ds.Rows)
            {
                //if (row.RowState != DataRowState.Modified)
                //{
                //    continue;
                //}
                Variable varr = new Variable();
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
                //Utlity.Log("varr.LOV:  " + varr.LOV, logFilePath);
                varr.AddVarToTemplate = Utlity.getBoolValue(row, "AddVarToTemplate", logFilePath);
                //Utlity.Log("varr.AddVarToTemplate:  " + varr.AddVarToTemplate, logFilePath);
                //varr.AddPartToTemplate = getValue(row, "AddPartToTemplate", logFilePath);
                varr.AddPartToTemplate = "Y";
                varr.variableType = Utlity.getValue(row, "variableType", logFilePath);
                varr.UnitType = Utlity.getValue(row, "UnitType", logFilePath);
                variablesList.Add(varr);
                //Utlity.Log(varr.systemName, logFilePath);
            }
            if (variablesList.Count == 0)
            {
                MessageBox.Show("No Variables to Update");
                return;
            }
            List<object> arguments = new List<object>();
            arguments.Add(assemFileName);
            arguments.Add(variablesList);
            arguments.Add(logFilePath);
            arguments.Add(partEnabledDictionary);
            if (backgroundWorker3.IsBusy != true)
            {
                MyCustomDialog4View.EnableDisableAllControls(false, this);
                backgroundWorker3.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker 3 is Busy, Sync SE Formula", logFilePath);
            }
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String assemblyFileName = (String)genericlist[0];
            List<Variable> variablesList = (List<Variable>)genericlist[1];
            String logFilePath = (String)genericlist[2];
            Dictionary<String, bool> partEnablementDictionary = (Dictionary<String, bool>)genericlist[3];
            Utlity.Log("SyncSolidEdgeFormula variablesList :" + variablesList.Count, logFilePath);
            {
                Utlity.Log("Calling SyncSolidEdgeFormula", logFilePath);
                backgroundWorker3.ReportProgress(10);
                SolidEdgeFormulaSync.SyncSolidEdgeFormula(assemblyFileName, variablesList, partEnablementDictionary, logFilePath);
                backgroundWorker3.ReportProgress(50);
                //variablesList.Clear();
                Utlity.Log("SyncSolidEdgeFormula Completed @ " + System.DateTime.Now.ToString(), logFilePath);
                SolidEdgeData1.readVariablesForEachOccurence(assemblyFileName, logFilePath);
                Utlity.Log("readVariablesForEachOccurence Completed @ " + System.DateTime.Now.ToString(), logFilePath);
                Utlity.Log("-----------------------------------------------------------------", logFilePath);
                backgroundWorker3.ReportProgress(95);

                genericlist[0] = assemblyFileName;
                genericlist[1] = variablesList;
                genericlist[2] = logFilePath;
                genericlist[3] = partEnablementDictionary;
                e.Result = genericlist;
                MessageBox.Show("Synced Formulae To Solid Edge Successfully: ");
                MyCustomDialog4View.EnableDisableAllControls(true, this);
            }

        }



        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = (List<object>)e.Result;
            if (genericlist == null || genericlist.Count == 0)
            {
                return;
            }

            String assemblyFileName = (String)genericlist[0];
            List<Variable> variablesList = (List<Variable>)genericlist[1];
            String logFilePath = (String)genericlist[2];
            Dictionary<String, bool> partEnablementDictionary = (Dictionary<String, bool>)genericlist[3];

            //List<Variable> ALLVariablesList = SolidEdgeData1.getVariableDetails();
            //Utlity.Log("variableList Size: " + ALLVariablesList.Count, logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            UpdateVariablesInManageVariablesTab(partEnablementDictionary, variablesList, logFilePath);   

        }

        private void UpdateVariablesInManageVariablesTab(Dictionary<String, bool> partEnablementDictionary, List<Variable> variablesList,String logFilePath)
        {
            this.progressBar2.Visible = false;
            allVarList.Clear();
            //allVarList = SolidEdgeData1.getVariableDetails();
            //MessageBox.Show(logFilePath);

            Dictionary<String, List<Variable>> UserVariableDictionary = utils.Utlity.BuildVariableDictionary(variablesList, "");
            if (UserVariableDictionary.Count == 0 || UserVariableDictionary == null)
            {
                Utlity.Log("No Variables in UI to Show", logFilePath);
                //MessageBox.Show("No Variables in UI to Show");
                return;
            }


            Dictionary<String, List<Variable>> variableDictionary = SolidEdgeData1.getVariablesDictionaryDetails();
            if (variableDictionary.Count == 0 || variableDictionary == null)
            {
                Utlity.Log("No Variables in SolidEdge to Show", logFilePath);
                //MessageBox.Show("No Variables in SolidEdge to Show");
                return;
            }
            if (partEnablementDictionary == null || partEnablementDictionary.Count == 0)
            {
                Utlity.Log("Data Missing to Show the UI: partEnabledDictionary", logFilePath);
                //MessageBox.Show("Data Missing to Show the UI: partEnabledDictionary");
                return;
            }
            int childCnt = partEnablementDictionary.Count;
            foreach (String s in partEnablementDictionary.Keys)
            {
                String occName = s;
                bool value = false;
                partEnablementDictionary.TryGetValue(s, out value);
                if (value == true)
                {
                    // SE Data -- Updated Variable Details (Formula Save Action)
                    List<Variable> variablesList1 = null;
                    bool Success = variableDictionary.TryGetValue(occName, out variablesList1);
                    
                    //MessageBox.Show(occName + ":::" + variablesList.Count.ToString());
                    if (Success == true)
                    {
                        // UI Data -- AddVarToTemplate & LOV
                        List<Variable> variablesList2 = null;
                        bool Success1 = UserVariableDictionary.TryGetValue(occName, out variablesList2);
                        //Utlity.Log(" UI variablesList: " + variablesList2.Count, logFilePath);
                        //Utlity.Log(" SE variablesList: " + variablesList1.Count, logFilePath);
                        variablesList1 = ViewUtils.MergeUserChangesToVariables(variablesList2, variablesList1, logFilePath);
                        if (variablesList1 != null && variablesList1.Count != 0)
                        {
                            allVarList.AddRange(variablesList1);
                        }
                    }
                }
            }

            List<Variable> newAllVarsList = PopulateSKYname(allVarList);
            if (newAllVarsList == null || newAllVarsList.Count == 0)
            {
                MessageBox.Show("No variables to show..");
                return;

            }


            DataTable table = null;
            table = Utlity.ConvertToDataTable(newAllVarsList);
            if (table == null)
            {
                MessageBox.Show("Unable to Convert Variables to Table");
                return;
            }

            //MessageBox.Show(table.Rows.Count.ToString());            
            this.dataGridView1.DataSource = table;

            if (UserVariableDictionary != null)
                UserVariableDictionary.Clear();
            if (allVarList != null)
                allVarList.Clear();
        }

        

        private void MyCustomDialog4_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            About a = new About();
            a.ShowDialog();
        }

        DataGridViewRow m_clickedRow = null;
        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Ignore if a column or row header is clicked
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                if (e.ColumnIndex == dataGridView1.Columns["AddVarToTemplate"].Index)
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        DataGridViewCell clickedCell = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex];
                        DataGridViewRow clickedRow = (sender as DataGridView).Rows[e.RowIndex];
                        m_clickedRow = clickedRow;

                        // Here you can do whatever you want with the cell
                        this.dataGridView1.CurrentCell = clickedCell;  // Select the clicked cell, for instance

                        // Get mouse position relative to the vehicles grid
                        var relativeMousePosition = dataGridView1.PointToClient(Cursor.Position);

                        // Show the context menu
                        this.contextMenuStrip2.Show(dataGridView1, relativeMousePosition);
                    }
                }
            }
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (m_clickedRow != null)
            {
                String partName = (String)m_clickedRow.Cells["PartName"].Value;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.DataBoundItem is Subro.Controls.GroupRow)
                    {
                        continue;
                    }
                    String currentRowPartName = (String)row.Cells["PartName"].Value;
                    if (currentRowPartName != null && currentRowPartName.Equals("") == false)
                    {
                        if (currentRowPartName.Equals(partName, StringComparison.OrdinalIgnoreCase) == true)
                        {
                            row.Cells["AddVarToTemplate"].Value = true;
                        }
                    }
                }
            }
            dataGridView1.Columns["PartName"].Visible = false;

        }

        private void unselectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (m_clickedRow != null)
            {
                String partName = (String)m_clickedRow.Cells["PartName"].Value;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.DataBoundItem is Subro.Controls.GroupRow)
                    {
                        continue;
                    }
                    String currentRowPartName = (String)row.Cells["PartName"].Value;
                    if (currentRowPartName != null && currentRowPartName.Equals("") == false)
                    {
                        if (currentRowPartName.Equals(partName, StringComparison.OrdinalIgnoreCase) == true)
                        {
                            row.Cells["AddVarToTemplate"].Value = false;
                        }
                    }
                }
            }
            dataGridView1.Columns["PartName"].Visible = false;

        }
        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (grouper != null) grouper.ExpandAll();
            if (chkSelectAll.CheckState == CheckState.Checked)
            {
                MessageBox.Show("All Part Variables Would be Selected: " + dataGridView1.Rows.Count);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.DataBoundItem is Subro.Controls.GroupRow)
                {
                    continue;
                }
                String currentRowPartName = (String)row.Cells["PartName"].Value;
                if (currentRowPartName != null && currentRowPartName.Equals("") == false)
                {
                    if (chkSelectAll.CheckState == CheckState.Checked)
                    {
                        row.Cells["AddVarToTemplate"].Value = true;
                    }
                    else
                    {
                        row.Cells["AddVarToTemplate"].Value = false;
                    }

                }
            }
            if (grouper != null) grouper.CollapseAll();

        }

        private void chkEditFormula_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEditFormula.CheckState == CheckState.Checked)
            {
                MessageBox.Show("Formula Property is Editable Now");
                this.dataGridView1.Columns["name"].ReadOnly = false;
                this.dataGridView1.Columns["Formula"].ReadOnly = false;

                this.dataGridView1.Columns["rangeLow"].ReadOnly = true;
                this.dataGridView1.Columns["rangeHigh"].ReadOnly = true;
                this.dataGridView1.Columns["LOV"].ReadOnly = true;
            }
            else if (chkEditFormula.CheckState == CheckState.Unchecked)
            {
                this.dataGridView1.Columns["name"].ReadOnly = true;
                this.dataGridView1.Columns["Formula"].ReadOnly = true;

                this.dataGridView1.Columns["rangeLow"].ReadOnly = false;
                this.dataGridView1.Columns["rangeHigh"].ReadOnly = false;
                this.dataGridView1.Columns["LOV"].ReadOnly = false;

            }
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Visible = true;
            this.progressBar1.Value = e.ProgressPercentage;

        }

        private void backgroundWorker3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar2.Visible = true;
            this.progressBar2.Value = e.ProgressPercentage;
        }

        private List<Variable> PopulateSKYname(List<Variable> allVariablesList)
        {
            Dictionary<String, String> SKYnameDictionary = SolidEdgeHighLighter.getSKYnameDictionary();
            if (SKYnameDictionary.Keys.Count == 0) return null;

            List<Variable> newAllVarsList = new List<Variable>();
            foreach (Variable var in allVariablesList)
            {
                String fileName = var.PartName;
                if (fileName != null && fileName.Equals("") == false)
                {
                    String SKYname = "";
                    SKYnameDictionary.TryGetValue(fileName, out SKYname);
                    var.Skyname = fileName + ";" + SKYname;
                    newAllVarsList.Add(var);
                }
                else
                {
                    // filename is empty, need filename to group
                }

            }
            return newAllVarsList;
        }

        private List<FeatureLine> PopulateSKYname(List<FeatureLine> allFeatureList)
        {
            //MessageBox.Show("allFeatureList Count:" + allFeatureList.Count);
            Dictionary<String, String> SKYnameDictionary = SolidEdgeHighLighter.getSKYnameDictionary();
            if (SKYnameDictionary.Keys.Count == 0) return null;
            //MessageBox.Show("SKYnameDictionary Count:" + SKYnameDictionary.Count);
            List<FeatureLine> featureList = new List<FeatureLine>();
            foreach (FeatureLine feature in allFeatureList)
            {
                String fileName = feature.PartName;
                if (fileName != null && fileName.Equals("") == false)
                {
                    String SKYname = "";
                    SKYnameDictionary.TryGetValue(fileName, out SKYname);
                    //MessageBox.Show("SKYname:" + SKYname);
                    feature.SKYname = fileName + ";" + SKYname;
                    featureList.Add(feature);
                }
                else
                {
                    // filename is empty, need filename to group
                }

            }
            return featureList;
        }

       

      

        
    }
}
