using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using ExcelLink.Forms;

namespace ExcelLink.Common
{
    public class ImportEventHandler : IExternalEventHandler
    {
        private string _excelFile;
        private Document _doc;
        private frmParaExport _form;

        public string GetName() => "Import Data from Excel";

        public void SetData(string excelFile, Document doc, frmParaExport form)
        {
            _excelFile = excelFile;
            _doc = doc;
            _form = form;
        }

        public void Execute(UIApplication app)
        {
            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;

            try
            {
                excel = new Excel.Application();
                workbook = excel.Workbooks.Open(_excelFile);

                worksheet = workbook.Worksheets.Cast<Excel.Worksheet>()
                                    .FirstOrDefault(s => s.Name != "Color Legend");

                if (worksheet == null)
                {
                    TaskDialog.Show("Error", "Could not find a valid worksheet to import from.");
                    return;
                }

                usedRange = worksheet.UsedRange;

                if (usedRange == null || usedRange.Rows.Count < 2)
                {
                    TaskDialog.Show("Error", "The selected worksheet is empty or does not contain any data rows.");
                    return;
                }

                List<string> headers = new List<string>();
                int firstDataColumn = 2;

                for (int j = firstDataColumn; j <= usedRange.Columns.Count; j++)
                {
                    var headerCell = usedRange.Cells[1, j] as Excel.Range;
                    if (headerCell != null && headerCell.Value2 != null)
                    {
                        headers.Add(headerCell.Value2.ToString());
                    }
                }

                List<string> errorMessages = new List<string>();

                using (Transaction t = new Transaction(_doc, "Import Parameters from Excel"))
                {
                    t.Start();
                    for (int i = 2; i <= usedRange.Rows.Count; i++)
                    {
                        var idCell = usedRange.Cells[i, 1] as Excel.Range;

                        if (idCell == null || idCell.Value2 == null) continue;

                        string idString = idCell.Value2.ToString();
                        int elementIdInt;

                        if (!int.TryParse(idString, out elementIdInt))
                        {
                            errorMessages.Add($"Row {i}: Failed to parse ElementId '{idString}'. Skipping row.");
                            continue;
                        }

                        ElementId elementId = new ElementId(elementIdInt);
                        Element element = _doc.GetElement(elementId);

                        if (element != null)
                        {
                            for (int j = 0; j < headers.Count; j++)
                            {
                                string paramName = headers[j];
                                var paramCell = usedRange.Cells[i, j + firstDataColumn] as Excel.Range;
                                string paramValue = paramCell?.Value2?.ToString();

                                if (!string.IsNullOrEmpty(paramValue))
                                {
                                    try
                                    {
                                        Parameter curParam = element.LookupParameter(paramName);
                                        if (curParam != null && curParam.IsReadOnly) continue;

                                        Utils.SetParameterValue(element, paramName, paramValue);
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Row {i}: Failed to set parameter '{paramName}' with value '{paramValue}'. Error: {ex.Message}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            errorMessages.Add($"Row {i}: Element with ID '{elementIdInt}' not found in model. Skipping row.");
                        }
                        _form.Dispatcher.Invoke(() => _form.UpdateProgressBar((int)((double)(i - 1) / (usedRange.Rows.Count - 1) * 100)));
                    }
                    t.Commit();
                }

                _form.Dispatcher.Invoke(() => _form.HideProgressBar());

                if (errorMessages.Any())
                {
                    string summary = "Import completed with errors:\n" + string.Join("\n", errorMessages.Take(10));
                    if (errorMessages.Count > 10)
                    {
                        summary += $"\n...and {errorMessages.Count - 10} more errors.";
                    }
                    TaskDialog.Show("Import Completed with Errors", summary);
                }
                else
                {
                    TaskDialog.Show("Success", "Import completed successfully!");
                }
            }
            catch (Exception ex)
            {
                _form.Dispatcher.Invoke(() => _form.HideProgressBar());
                TaskDialog.Show("Error", $"Failed to import parameters:\n{ex.Message}");
            }
            finally
            {
                if (workbook != null) workbook.Close(false);
                if (excel != null) excel.Quit();
                if (usedRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }
    }
}