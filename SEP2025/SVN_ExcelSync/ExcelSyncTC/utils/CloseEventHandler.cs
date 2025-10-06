using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSyncTC.utils
{
    class CloseEventHandler
    {
        static private Excel.Application app;
        static private Excel.Workbook workbook;
        static private bool exit = false;

        public CloseEventHandler(Excel.Application application)
        {
            app = application;
            if (app != null) workbook = app.ActiveWorkbook;
            //if (workbook!=null) worksheet = workbook.Worksheets.get_Item(1) as Excel.Worksheet;
            if (workbook == null) return;

            app.WorkbookBeforeClose +=
              new Excel.AppEvents_WorkbookBeforeCloseEventHandler(
              App_WorkbookBeforeClose);

            workbook.BeforeClose +=
              new Excel.WorkbookEvents_BeforeCloseEventHandler(
              Workbook_BeforeClose);

            //app.WorkbookBeforePrint +=
            //  new Excel.AppEvents_WorkbookBeforePrintEventHandler(
            //  App_WorkbookBeforePrint);

            //workbook.BeforePrint +=
            //  new Excel.WorkbookEvents_BeforePrintEventHandler(
            //  Workbook_BeforePrint);

            //app.WorkbookBeforeSave +=
            //  new Excel.AppEvents_WorkbookBeforeSaveEventHandler(
            //  App_WorkbookBeforeSave);

            //workbook.BeforeSave +=
            //  new Excel.WorkbookEvents_BeforeSaveEventHandler(
            //  Workbook_BeforeSave);

            //app.WorkbookOpen +=
            //  new Excel.AppEvents_WorkbookOpenEventHandler(
            //  App_WorkbookOpen);

            while (exit == false)
                System.Windows.Forms.Application.DoEvents();

            //app.Quit();
        }

        static void App_WorkbookBeforeClose(Excel.Workbook workbook,
          ref bool cancel)
        {
            Console.WriteLine(String.Format(
              "Application.WorkbookBeforeClose({0})",
              workbook.Name));
            utils.Utlity.ModSheetsInSession.Clear();
        }

        static void Workbook_BeforeClose(ref bool cancel)
        {
            Console.WriteLine("Workbook.BeforeClose()");
            exit = true;
            utils.Utlity.ModSheetsInSession.Clear();
        }

        static void App_WorkbookBeforePrint(Excel.Workbook workbook,
          ref bool cancel)
        {
            Console.WriteLine(String.Format(
              "Application.WorkbookBeforePrint({0})",
              workbook.Name));
            cancel = true; // Don't allow printing
        }

        static void Workbook_BeforePrint(ref bool cancel)
        {
            Console.WriteLine("Workbook.BeforePrint()");
            cancel = true; // Don't allow printing
        }

        static void App_WorkbookBeforeSave(Excel.Workbook workbook,
          bool saveAsUI, ref bool cancel)
        {
            Console.WriteLine(String.Format(
              "Application.WorkbookBeforeSave({0},{1})",
              workbook.Name, saveAsUI));
            cancel = true; // Don't allow saving
        }

        static void Workbook_BeforeSave(bool saveAsUI, ref bool cancel)
        {
            Console.WriteLine(String.Format(
              "Workbook.BeforePrint({0})",
              saveAsUI));
            cancel = true; // Don't allow saving
        }

        static void App_WorkbookOpen(Excel.Workbook workbook)
        {
            Console.WriteLine(String.Format(
              "Appplication.WorkbookOpen({0})",
              workbook.Name));
        }
    }
}
