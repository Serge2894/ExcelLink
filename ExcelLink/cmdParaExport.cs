using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelLink.Forms;
using ExcelLink.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;
using System.Drawing;
using System.Text;

namespace ExcelLink
{
    [Transaction(TransactionMode.Manual)]
    public class cmdParaExport : IExternalCommand
    {
        public static ExternalEvent ImportExternalEvent;
        public static ImportEventHandler ImportEventHandler;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                if (ImportExternalEvent == null)
                {
                    ImportEventHandler = new ImportEventHandler();
                    ImportExternalEvent = ExternalEvent.Create(ImportEventHandler);
                }

                frmParaExport form = new frmParaExport(doc, ImportExternalEvent, ImportEventHandler);
                form.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnParaExport";
            string buttonTitle = "Para\rExport";

            ButtonDataClass myButtonData = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Export/Import parameters to/from Excel");

            return myButtonData.Data;
        }
    }
}