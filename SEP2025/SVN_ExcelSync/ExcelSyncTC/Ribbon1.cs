using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ExcelSyncTC.controller;
using System.Globalization;
using ExcelSyncTC.utils;
using ExcelSyncTC.opInterfaces;
using System.IO;
using System.Windows.Forms;
using System.ComponentModel;

namespace ExcelSyncTC
{
    public partial class Ribbon1
    {
       

      
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
            //// 29 Sept, Reference - http://blogs.infoextract.in/excel-events-using-c-dot-net/
            //// 15 Oct,Reference - Add Ins Not Loaded Some times -- https://stackoverflow.com/questions/13936388/my-excel-2010-add-in-only-shows-up-when-opening-a-blank-workbook-wont-show-up
            try
            {
                Globals.ThisAddIn.Application.SheetChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler(Application_SheetChange);
                Globals.ThisAddIn.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

                Globals.ThisAddIn.Application.SheetCalculate += new Microsoft.Office.Interop.Excel.AppEvents_SheetCalculateEventHandler(Application_SheetCalculate);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //try
            //{
            //    ChangeEventHandler ch = new ChangeEventHandler(Globals.ThisAddIn.Application);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("ChangeEventHandler: " + ex.Message);
            //}

            //try
            //{
            //    CloseEventHandler ch = new CloseEventHandler(Globals.ThisAddIn.Application);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("CloseEventHandler: " + ex.Message);
            //}
        

        }

        private void Application_SheetCalculate(object Sh)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)Sh;
                //MessageBox.Show("Application_SheetCalculate: " + sheet.Name);
                if (utils.Utlity.ModSheetsInSession.Contains(sheet.Name) == false)
                {
                    utils.Utlity.ModSheetsInSession.Add(sheet.Name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Application_SheetCalculate: " + ex.Message);
            }
            
        }
        void Application_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
        {
            utils.Utlity.ModSheetsInSession.Clear();

        }

        void Application_SheetChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)Sh;
                //MessageBox.Show("Application_SheetChange: " + sheet.Name);
                if (utils.Utlity.ModSheetsInSession.Contains(sheet.Name) == false)
                {
                    utils.Utlity.ModSheetsInSession.Add(sheet.Name);

                }                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Application_SheetChange: " + ex.Message);
            }
            //MessageBox.Show(sheet.Name);
            //string changedRange = Target.get_Address(
            //  Excel.XlReferenceStyle.xlA1);

        }

        // Sync TE (3D Only) - Simone Request - 04 December
        public static ExcelSyncDialog dialog = new ExcelSyncDialog();
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {            

            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;
            if (dialog.IsDisposed == false)
                dialog.Show();
            else
            {
                dialog = new ExcelSyncDialog();
                dialog.Show();
            }
            
            
        }        

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            xlApp = null;

        }

        // Sync TE (DWG Only) - Simone Request - 04 December
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = Globals.ThisAddIn.Application;
            SyncDwg dwgSync = new SyncDwg();
            dwgSync.Show();
        }
    }
}
