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
                // Show progress bar
                _form.Dispatcher.Invoke(() => _form.ShowProgressBar());

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
                int firstDataColumn = 2; // Skip the first column (Element ID)

                // Parse headers - extract just the parameter name from multi-line headers
                for (int j = firstDataColumn; j <= usedRange.Columns.Count; j++)
                {
                    var headerCell = usedRange.Cells[1, j] as Excel.Range;
                    if (headerCell != null && headerCell.Value2 != null)
                    {
                        string fullHeader = headerCell.Value2.ToString();
                        // Extract just the parameter name (first line)
                        string paramName = fullHeader.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                        headers.Add(paramName);
                    }
                }

                List<string> errorMessages = new List<string>();
                List<string> successMessages = new List<string>();
                int totalRows = usedRange.Rows.Count - 1; // Exclude header row
                int processedRows = 0;

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
                            errorMessages.Add($"Row {i}: Invalid ElementId '{idString}'");
                            continue;
                        }

                        ElementId elementId = new ElementId(elementIdInt);
                        Element element = _doc.GetElement(elementId);

                        if (element != null)
                        {
                            int updatedParams = 0;

                            for (int j = 0; j < headers.Count; j++)
                            {
                                string paramName = headers[j];
                                var paramCell = usedRange.Cells[i, j + firstDataColumn] as Excel.Range;

                                // Skip if cell is null or empty
                                if (paramCell == null || paramCell.Value2 == null) continue;

                                string paramValue = paramCell.Value2.ToString();

                                // Skip if the cell has grey background (parameter doesn't exist)
                                var cellColor = paramCell.Interior.Color;
                                if (cellColor != null)
                                {
                                    int colorValue = Convert.ToInt32(cellColor);
                                    // Check if it's grey color (D3D3D3)
                                    if (colorValue == ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3")))
                                    {
                                        continue;
                                    }
                                }

                                try
                                {
                                    // First try to get instance parameter
                                    Parameter param = element.LookupParameter(paramName);
                                    Element targetElement = element;

                                    // If not found, try type parameter
                                    if (param == null)
                                    {
                                        Element typeElem = _doc.GetElement(element.GetTypeId());
                                        if (typeElem != null)
                                        {
                                            param = typeElem.LookupParameter(paramName);
                                            targetElement = typeElem;
                                        }
                                    }

                                    // If still not found, try built-in parameter
                                    if (param == null)
                                    {
                                        BuiltInParameter bip = Utils.GetBuiltInParameterByName(paramName);
                                        if (bip != BuiltInParameter.INVALID)
                                        {
                                            param = element.get_Parameter(bip);
                                            targetElement = element;

                                            if (param == null)
                                            {
                                                Element typeElem = _doc.GetElement(element.GetTypeId());
                                                if (typeElem != null)
                                                {
                                                    param = typeElem.get_Parameter(bip);
                                                    targetElement = typeElem;
                                                }
                                            }
                                        }
                                    }

                                    if (param != null && !param.IsReadOnly)
                                    {
                                        // Get current value to check if it's different
                                        string currentValue = Utils.GetParameterValue(targetElement, paramName);

                                        if (currentValue != paramValue)
                                        {
                                            bool success = Utils.SetParameterValue(targetElement, paramName, paramValue);
                                            if (success)
                                            {
                                                updatedParams++;
                                            }
                                            else
                                            {
                                                errorMessages.Add($"Row {i}: Failed to set parameter '{paramName}' to '{paramValue}'");
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorMessages.Add($"Row {i}: Error with parameter '{paramName}': {ex.Message}");
                                }
                            }

                            if (updatedParams > 0)
                            {
                                successMessages.Add($"Element {elementIdInt}: Updated {updatedParams} parameter(s)");
                            }
                        }
                        else
                        {
                            errorMessages.Add($"Row {i}: Element ID {elementIdInt} not found in model");
                        }

                        processedRows++;
                        _form.Dispatcher.Invoke(() =>
                            _form.UpdateProgressBar((int)((double)processedRows / totalRows * 100)));
                    }

                    t.Commit();
                }

                _form.Dispatcher.Invoke(() => _form.HideProgressBar());

                // Show results
                string message = "";

                if (successMessages.Any())
                {
                    message = $"Import completed successfully!\n\n";
                    message += $"Updated elements: {successMessages.Count}\n";

                    if (successMessages.Count <= 10)
                    {
                        message += "\nDetails:\n" + string.Join("\n", successMessages);
                    }
                    else
                    {
                        message += "\nDetails:\n" + string.Join("\n", successMessages.Take(5));
                        message += $"\n... and {successMessages.Count - 5} more elements";
                    }
                }

                if (errorMessages.Any())
                {
                    if (string.IsNullOrEmpty(message))
                    {
                        message = "Import completed with errors:\n\n";
                    }
                    else
                    {
                        message += "\n\nErrors encountered:\n";
                    }

                    message += string.Join("\n", errorMessages.Take(10));
                    if (errorMessages.Count > 10)
                    {
                        message += $"\n... and {errorMessages.Count - 10} more errors";
                    }
                }

                if (string.IsNullOrEmpty(message))
                {
                    message = "Import completed. No parameters were updated (all values may already be up to date).";
                }

                TaskDialog.Show("Import Results", message);
            }
            catch (Exception ex)
            {
                _form.Dispatcher.Invoke(() => _form.HideProgressBar());
                TaskDialog.Show("Error", $"Failed to import parameters:\n{ex.Message}");
            }
            finally
            {
                // Clean up Excel objects
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excel != null)
                {
                    excel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                }
            }
        }
    }
}