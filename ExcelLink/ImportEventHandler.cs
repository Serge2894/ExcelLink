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

                List<ImportErrorItem> errorMessages = new List<ImportErrorItem>();
                int totalRows = usedRange.Rows.Count - 1; // Exclude header row
                int processedRows = 0;
                int updatedElementsCount = 0;

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
                            errorMessages.Add(new ImportErrorItem { ElementId = idString, Description = $"Invalid ElementId '{idString}'" });
                            continue;
                        }

                        ElementId elementId = new ElementId(elementIdInt);
                        Element element = _doc.GetElement(elementId);

                        if (element != null)
                        {
                            bool elementUpdated = false;

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
                                    Parameter param = element.LookupParameter(paramName);
                                    Element targetElement = element;

                                    if (param == null)
                                    {
                                        Element typeElem = _doc.GetElement(element.GetTypeId());
                                        if (typeElem != null)
                                        {
                                            param = typeElem.LookupParameter(paramName);
                                            targetElement = typeElem;
                                        }
                                    }

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
                                        string currentValue = Utils.GetParameterValue(targetElement, paramName);

                                        if (currentValue != paramValue)
                                        {
                                            bool success = Utils.SetParameterValue(targetElement, paramName, paramValue);
                                            if (success)
                                            {
                                                elementUpdated = true;
                                            }
                                            else
                                            {
                                                errorMessages.Add(new ImportErrorItem { ElementId = idString, Description = $"Error updating parameter '{paramName}' with value '{paramValue}'" });
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorMessages.Add(new ImportErrorItem { ElementId = idString, Description = $"Error updating parameter '{paramName}': {ex.Message}" });
                                }
                            }
                            if (elementUpdated)
                            {
                                updatedElementsCount++;
                            }
                        }
                        else
                        {
                            errorMessages.Add(new ImportErrorItem { ElementId = idString, Description = "Element ID not found in model" });
                        }

                        processedRows++;
                        _form.Dispatcher.Invoke(() =>
                            _form.UpdateProgressBar((int)((double)processedRows / totalRows * 100)));
                    }

                    t.Commit();
                }

                // Keep the progress bar at 100% and green after completion
                _form.Dispatcher.Invoke(() => _form.UpdateProgressBar(100));

                if (errorMessages.Any())
                {
                    // Show custom form with errors
                    _form.Dispatcher.Invoke(() =>
                    {
                        var failForm = new frmImportFailed(errorMessages);
                        failForm.ShowDialog();
                    });
                }
                else if (updatedElementsCount > 0)
                {
                    // Show info dialog for success as requested
                    _form.Dispatcher.Invoke(() =>
                    {
                        frmInfoDialog infoDialog = new frmInfoDialog("Model updated successfully");
                        infoDialog.ShowDialog();
                    });
                }
                else
                {
                    // Show info dialog if no parameters were updated
                    _form.Dispatcher.Invoke(() =>
                    {
                        frmInfoDialog infoDialog = new frmInfoDialog("Import completed. \nNo parameters were updated.");
                        infoDialog.ShowDialog();
                    });
                }
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