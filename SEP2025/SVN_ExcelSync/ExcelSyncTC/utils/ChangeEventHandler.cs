using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSyncTC.utils
{
    public class ChangeEventHandler
    {
        private Excel.Application app;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        object missing = System.Type.Missing;

        public ChangeEventHandler(Excel.Application application)
        {
            this.app = application;
            if (app!= null) workbook = app.ActiveWorkbook;
            //if (workbook!=null) worksheet = workbook.Worksheets.get_Item(1) as Excel.Worksheet;
            if (workbook == null) return;

            app.SheetChange +=
              new Excel.AppEvents_SheetChangeEventHandler(
              App_SheetChange);

            workbook.SheetChange +=
              new Excel.WorkbookEvents_SheetChangeEventHandler(
              Workbook_SheetChange);

            //worksheet.Change +=
            //  new Excel.DocEvents_ChangeEventHandler(
            //  Worksheet_Change);
        }

        // Change events only pass worksheets, never charts.
        private string SheetName(object sheet)
        {
            Excel.Worksheet worksheet = sheet as Excel.Worksheet;
            return worksheet.Name;
        }

        private string RangeAddress(Excel.Range target)
        {
            return target.get_Address(missing, missing,
              Excel.XlReferenceStyle.xlA1, missing, missing);
        }

        void App_SheetChange(object sheet, Excel.Range target)
        {
            Console.WriteLine(String.Format(
              "Application.SheetChange({0},{1})",
              SheetName(sheet), RangeAddress(target)));

            if (utils.Utlity.ModSheetsInSession.Contains(SheetName(sheet)) == false)
            {
                utils.Utlity.ModSheetsInSession.Add(SheetName(sheet));

            }
            //System.Windows.Forms.MessageBox.Show(SheetName(sheet));
        }

        void Workbook_SheetChange(object sheet, Excel.Range target)
        {
            Console.WriteLine(String.Format(
              "Workbook.SheetChange({0},{1})",
              SheetName(sheet), RangeAddress(target)));

            if (utils.Utlity.ModSheetsInSession.Contains(SheetName(sheet)) == false)
            {
                utils.Utlity.ModSheetsInSession.Add(SheetName(sheet));

            }
            //System.Windows.Forms.MessageBox.Show(SheetName(sheet));
        }

        //void Worksheet_Change(Excel.Range target)
        //{
        //    Console.WriteLine(String.Format(
        //      "Worksheet.Change({0})",
        //      RangeAddress(target)));
        //}
    }
}
