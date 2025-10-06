using DemoAddInTC.controller;
using DemoAddInTC.model;
using DemoAddInTC.opInterfaces;
using DemoAddInTC.utils;
using DgvFilterPopup;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC
{
    public partial class MyCustomDialog1 : Form
    {
        Subro.Controls.DataGridViewGrouper grouper = null;

        public MyCustomDialog1()
        {
            InitializeComponent();
            populateDataFromSE();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }


        private void populateDataFromSE()
        {
            List<Variable> allVarList = new List<Variable>();
            allVarList = ExcelData.getVariableDetails();
            DataTable table = null;
            table = Utlity.ConvertToDataTable(allVarList);

            //MessageBox.Show(table.Rows.Count.ToString());
            this.dataGridView1.DataSource = table;
            DgvFilterManager filterManager = new DgvFilterManager(dataGridView1);
            try
            {
                grouper = new Subro.Controls.DataGridViewGrouper(dataGridView1);
                grouper.SetGroupOn<Variable>(t => t.PartName);
                grouper.SetGroupOn(this.dataGridView1.Columns["PartName"]);
                //grouper.Options.StartCollapsed = true;
                grouper.CollapseAll();
                grouper.DisplayGroup += grouper_DisplayGroup;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            hideColumns();
            ApplyAdditionalSettings();
        }

        //optionally, you can customize the grouping display by subscribing to the DisplayGroup event
        void grouper_DisplayGroup(object sender, Subro.Controls.GroupDisplayEventArgs e)
        {
            e.BackColor = (e.Group.GroupIndex % 2) == 0 ? Color.LightPink : Color.LightSkyBlue;            
            //e.Header = "[" + e.Header + "], grp: " + e.Group.GroupIndex;
            e.Header = "";
            e.DisplayValue = e.DisplayValue;
            e.Summary = "(" + e.Group.Count + " )";
        }

        private void ApplyAdditionalSettings()
        {
            this.dataGridView1.Columns["PartName"].ReadOnly = true;
            this.dataGridView1.Columns["name"].ReadOnly = true;
            this.dataGridView1.Columns["systemName"].ReadOnly = true;
            this.dataGridView1.Columns["unit"].ReadOnly = true;
            this.dataGridView1.Columns["rangeCondition"].ReadOnly = true;
            this.dataGridView1.Columns["Formula"].ReadOnly = true;

            this.dataGridView1.Columns["name"].Frozen = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Azure;
            //dataGridView1.DefaultCellStyle.SelectionBackColor = Color.White;
            this.dataGridView1.Columns["PartName"].Visible = false;
        }

        private void hideColumns()
        {
            this.dataGridView1.Columns["PartName"].Visible = true;
            this.dataGridView1.Columns["systemName"].Visible = false;
            this.dataGridView1.Columns["unit"].Visible = false;
            this.dataGridView1.Columns["rangeCondition"].Visible = false;

            this.dataGridView1.Columns["variableType"].Visible = true;
            this.dataGridView1.Columns["AddVarToTemplate"].Visible = true;
            this.dataGridView1.Columns["AddPartToTemplate"].Visible = true;
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            ClipboardUtils.OnDataGridViewPaste(sender, e);
        }

        // If User enters Y/N -- cascade the same to ALL cells for the corresponding Part.
        private void DataGridView1_CellValueChanged(
    object sender, DataGridViewCellEventArgs e)
        {
            //if (codeIsChanging == true)
            //{
            //    return;
            //}
            //MessageBox.Show(dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name);
            if (dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name == "AddPartToTemplate")
            {

                String partSelectedOrNot = (String)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["AddPartToTemplate"].Value;
                //MessageBox.Show(partSelectedOrNot);
                if (partSelectedOrNot.Equals("Y", StringComparison.OrdinalIgnoreCase) == true ||
                    partSelectedOrNot.Equals("N", StringComparison.OrdinalIgnoreCase) == true)
                {
                    String partName = (String)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells["PartName"].Value;
                    //MessageBox.Show(partName);
                    //codeIsChanging = true;
                    populatePartNameToALLRows(partName, partSelectedOrNot);
                    //codeIsChanging = false;
                }
                else
                {
                    MessageBox.Show("Enter Either Y/N Only to Include Part to Template");
                }
            }
        }

        private void populatePartNameToALLRows(string partName, String partSelectedOrNot)
        {
            this.dataGridView1.CellValueChanged -= this.DataGridView1_CellValueChanged;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["partName"].Value != null)
                {
                    if (row.DataBoundItem is Subro.Controls.GroupRow)
                    {
                        continue;
                    }
                    if (Convert.IsDBNull(row.Cells["partName"].Value) == false)
                    {
                        String cellPartName = (String)row.Cells["partName"].Value;
                        if (cellPartName.Equals(partName, StringComparison.OrdinalIgnoreCase))
                        {
                            row.Cells["AddPartToTemplate"].Value = partSelectedOrNot;
                            row.Cells["AddVarToTemplate"].Value = partSelectedOrNot;
                        }
                    }
                }
            }
            this.dataGridView1.CellValueChanged += this.DataGridView1_CellValueChanged;
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {



            //if (dataGridView1.SelectedCells.Count > 0 && e.RowIndex != -1 && e.Button == MouseButtons.Right)
            //{
            //    MessageBox.Show(dataGridView1.SelectedCells.Count.ToString());
            //    MessageBox.Show(e.RowIndex.ToString());
            //    dataGridView1.ContextMenuStrip = contextMenuStrip1;
            //}
            //else
            //{
            //    if (dataGridView1.ContextMenuStrip != null)
            //    {
            //        dataGridView1.ContextMenuStrip.Enabled = false;
            //        dataGridView1.ContextMenuStrip = null;
            //    }
            //}
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            MessageBox.Show("dataGridView1_KeyDown");
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Copy/paste operation failed. " + ex.Message, "Copy/Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CopyToClipboard()
        {
            //Copy to clipboard
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void PasteClipboardValue()
        {
            //Show Error if no cell is selected
            if (dataGridView1.SelectedCells.Count == 0)
            {
                MessageBox.Show("Please select a cell", "Paste", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Get the satring Cell
            DataGridViewCell startCell = GetStartCell(dataGridView1);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue = ClipBoardValues(Clipboard.GetText());

            int iRowIndex = startCell.RowIndex;
            foreach (int rowKey in cbValue.Keys)
            {
                int iColIndex = startCell.ColumnIndex;
                foreach (int cellKey in cbValue[rowKey].Keys)
                {
                    //Check if the index is with in the limit
                    if (iColIndex <= dataGridView1.Columns.Count - 1 && iRowIndex <= dataGridView1.Rows.Count - 1)
                    {
                        DataGridViewCell cell = dataGridView1[iColIndex, iRowIndex];

                        //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                        if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                            (!chkPasteToSelectedCells.Checked))
                            cell.Value = cbValue[rowKey][cellKey];
                    }
                    iColIndex++;
                }
                iRowIndex++;
            }
        }

        private DataGridViewCell GetStartCell(DataGridView dgView)
        {
            //get the smallest row,column index
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }

        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>> copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                //if an empty cell value copied, then set the dictionay with an empty string
                //else Set value to dictionary
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }

        private String getValue(DataGridViewRow row, String propName, String logFilePath)
        {
            String value = "";
            try
            {
                value = (String)row.Cells[propName].Value;
            }
            catch (Exception ex)
            {
                //Utlity.Log(ex.Message + ":::" + propName, logFilePath);
                return "";
            }
            return value;
        }

        private String getValue(DataRow row, String propName, String logFilePath)
        {
            String value = "";
            try
            {
                value = (String)row[propName];
            }
            catch (Exception ex)
            {
                //Utlity.Log(ex.Message + ":::" + propName, logFilePath);
                return "";
            }
            return value;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> genericlist = e.Argument as List<object>;
            String outputXLfileName = (String)genericlist[0];
            List<Variable> variablesList = (List<Variable>)genericlist[1];
            String logFilePath = (String)genericlist[2];
            //ExcelInterface.SaveDELTAToXL(outputXLfileName, variablesList,null, logFilePath);
            genericlist[0] = outputXLfileName;
            genericlist[1] = variablesList;
            genericlist[2] = logFilePath;
            e.Result = genericlist;

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> genericlist = (List<object>)e.Result;
            if (genericlist == null || genericlist.Count == 0)
            {
                return;
            }
            String outputXLfileName = (String)genericlist[0];
            List<Variable> variablesList = (List<Variable>)genericlist[1];
            String logFilePath = (String)genericlist[2];

            MessageBox.Show("Saved the Template Successfully: " + outputXLfileName);
            Utlity.Log("VariablesList Count:" + variablesList.Count, logFilePath);
            Utlity.Log("Utility Ended @ " + System.DateTime.Now.ToString(), logFilePath);
            Utlity.Log("-----------------------------------------------------------------", logFilePath);
            //this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Save Clicked");

            this.Dispose();
            String assemFileName = SolidEdgeData.getAssemblyFileName();
            if (assemFileName == null || assemFileName.Equals("") == true)
            {
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
            DataTable ds = (DataTable)grouper.GroupingSource.DataSource;
            Utlity.Log(dataGridView1.Rows.Count.ToString(), logFilePath);  
            List<Variable> variablesList = new List<Variable>();
            foreach (DataRow row in ds.Rows)
            {
                
                //if (row.IsNewRow == true) continue;
                //if (row.DataBoundItem is Subro.Controls.GroupRow)
                //{
                //    continue;
                //}
                Utlity.Log("Reading : " + (String)row["PartName"].ToString(), logFilePath);
                Variable varr = new Variable();
                varr.PartName = getValue(row, "PartName", logFilePath);
                varr.name = getValue(row, "name", logFilePath);
                varr.systemName = getValue(row, "systemName", logFilePath);
                varr.value = getValue(row, "value", logFilePath);
                varr.unit = getValue(row, "unit", logFilePath);
                varr.rangeLow = getValue(row, "rangeLow", logFilePath);
                varr.rangeHigh = getValue(row, "rangeHigh", logFilePath);
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
                varr.Formula = getValue(row, "Formula", logFilePath);
                varr.LOV = getValue(row, "LOV", logFilePath);
                //varr.AddVarToTemplate = getValue(row, "AddVarToTemplate", logFilePath);
                varr.AddPartToTemplate = getValue(row, "AddPartToTemplate", logFilePath);
                varr.variableType = getValue(row, "variableType", logFilePath);
                variablesList.Add(varr);
            }


            List<object> arguments = new List<object>();
            arguments.Add(outputXLfileName);
            arguments.Add(variablesList);
            arguments.Add(logFilePath);
            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync(arguments);
            }
            else
            {
                Utlity.Log("Background Worker is Busy, Save in CTC", logFilePath);
            }


        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //e.CellStyle.BackColor = Color.White;

            //if (dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name == "variableType")
            //{
            //    TextBox t = e.Control as TextBox;

            //    if (t != null)
            //    {
            //        t.AutoCompleteMode = AutoCompleteMode.Suggest;
            //        t.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //        AutoCompleteStringCollection data = new AutoCompleteStringCollection();
            //        data.AddRange(new String[] {"Limit", "LOV"});
            //        t.AutoCompleteCustomSource = data;
            //    }
            //}
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }
    }
}
