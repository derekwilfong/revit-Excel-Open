#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;


using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace RevitAddin_OpenExcel
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        static void openfile()
        {
            string path = @"C:";
            var xlApp = new Excel.Application();
            xlApp.Visible = true;

            Excel.Workbooks wb = xlApp.Workbooks;
            Excel.Workbook ws = wb.Open(path);

        }


        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            openfile();

            //using (Transaction tx = new Transaction(doc))
            //{
            //    tx.Start("Transaction Name");
            //    tx.Commit();
            //}

            return Result.Succeeded;
        }
    }
}
