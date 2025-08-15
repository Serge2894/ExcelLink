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

        public List<ImportErrorItem> ErrorMessages { get; private set; }
        public int UpdatedElementsCount { get; private set; }

        public string GetName() => "Import Data from Excel";

        public void SetData(string excelFile, Document doc, frmParaExport form)
        {
            _excelFile = excelFile;
            _doc = doc;
            _form = form;
            ErrorMessages = new List<ImportErrorItem>();
            UpdatedElementsCount = 0;
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
                    _form.Dispatcher.Invoke(() => _form.HideProgressBar());
                    return;
                }

                usedRange = worksheet.UsedRange;

                if (usedRange == null || usedRange.Rows.Count < 2)
                {
                    TaskDialog.Show("Error", "The selected worksheet is empty or does not contain any data rows.");
                    _form.Dispatcher.Invoke(() => _form.HideProgressBar());
                    return;
                }

                var headers = new List<(string Name, bool IsType)>();
                int firstDataColumn = 2; // Skip the first column (Element ID)

                for (int j = firstDataColumn; j <= usedRange.Columns.Count; j++)
                {
                    var headerCell = usedRange.Cells[1, j] as Excel.Range;
                    if (headerCell != null && headerCell.Value2 != null)
                    {
                        string fullHeader = headerCell.Value2.ToString();
                        var headerLines = fullHeader.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                        string paramName = headerLines.Length > 0 ? headerLines[0].Trim() : string.Empty;
                        bool isType = headerLines.Length > 1 && headerLines[1].Trim() == "(Type Parameter)";
                        headers.Add((paramName, isType));
                    }
                }

                var typeParamData = new Dictionary<Tuple<ElementId, string>, List<Tuple<string, string>>>();

                // Pre-pass to gather all type parameter values
                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    var idCell = usedRange.Cells[i, 1] as Excel.Range;
                    if (idCell == null || idCell.Value2 == null) continue;
                    string idString = idCell.Value2.ToString();
                    if (!int.TryParse(idString, out int elementIdInt)) continue;

                    ElementId elementId = new ElementId(elementIdInt);
                    Element element = _doc.GetElement(elementId);

                    if (element != null)
                    {
                        Element typeElem = _doc.GetElement(element.GetTypeId());
                        if (typeElem != null)
                        {
                            for (int j = 0; j < headers.Count; j++)
                            {
                                var (paramName, isType) = headers[j];
                                if (isType)
                                {
                                    var paramCell = usedRange.Cells[i, j + firstDataColumn] as Excel.Range;
                                    if (paramCell != null && paramCell.Value2 != null)
                                    {
                                        string paramValue = paramCell.Value2.ToString();
                                        var key = Tuple.Create(typeElem.Id, paramName);
                                        if (!typeParamData.ContainsKey(key))
                                        {
                                            typeParamData[key] = new List<Tuple<string, string>>();
                                        }
                                        typeParamData[key].Add(Tuple.Create(paramValue, idString));
                                    }
                                }
                            }
                        }
                    }
                }

                var inconsistentTypeParameters = new HashSet<Tuple<ElementId, string>>();
                foreach (var kvp in typeParamData)
                {
                    var groups = kvp.Value.GroupBy(t => t.Item1).ToList();
                    if (groups.Count > 1)
                    {
                        inconsistentTypeParameters.Add(kvp.Key);
                        foreach (var group in groups)
                        {
                            foreach (var item in group)
                            {
                                ErrorMessages.Add(new ImportErrorItem { ElementId = item.Item2, Description = $"Inconsistent value '{item.Item1}' for type parameter '{kvp.Key.Item2}'. Type parameters with the same ID must be filled the same." });
                            }
                        }
                    }
                }

                int totalRows = usedRange.Rows.Count - 1;
                int processedRows = 0;
                var updatedTypeParameters = new HashSet<Tuple<ElementId, string>>();

                using (Transaction t = new Transaction(_doc, "Import Parameters from Excel"))
                {
                    t.Start();

                    for (int i = 2; i <= usedRange.Rows.Count; i++)
                    {
                        var idCell = usedRange.Cells[i, 1] as Excel.Range;
                        if (idCell == null || idCell.Value2 == null) continue;
                        string idString = idCell.Value2.ToString();
                        if (!int.TryParse(idString, out int elementIdInt))
                        {
                            // Error already logged in pre-pass logic if needed
                            continue;
                        }

                        ElementId elementId = new ElementId(elementIdInt);
                        Element element = _doc.GetElement(elementId);

                        if (element != null)
                        {
                            bool elementUpdated = false;
                            Element typeElem = _doc.GetElement(element.GetTypeId());

                            for (int j = 0; j < headers.Count; j++)
                            {
                                var (paramName, isType) = headers[j];
                                var paramCell = usedRange.Cells[i, j + firstDataColumn] as Excel.Range;

                                if (paramCell == null || paramCell.Value2 == null) continue;
                                string paramValue = paramCell.Value2.ToString();
                                var cellColor = paramCell.Interior.Color;
                                if (cellColor != null)
                                {
                                    int colorValue = Convert.ToInt32(cellColor);
                                    if (colorValue == ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3")))
                                        continue;
                                }

                                if (isType && typeElem != null)
                                {
                                    var key = Tuple.Create(typeElem.Id, paramName);
                                    if (inconsistentTypeParameters.Contains(key))
                                    {
                                        continue; // Skip update for this inconsistent parameter
                                    }

                                    if (updatedTypeParameters.Contains(key))
                                    {
                                        continue; // Already updated this type parameter
                                    }
                                }

                                Parameter param = element.LookupParameter(paramName);
                                Element targetElement = element;

                                if (param == null && typeElem != null)
                                {
                                    param = typeElem.LookupParameter(paramName);
                                    targetElement = typeElem;
                                }

                                if (param == null)
                                {
                                    BuiltInParameter bip = Utils.GetBuiltInParameterByName(paramName);
                                    if (bip != BuiltInParameter.INVALID)
                                    {
                                        param = element.get_Parameter(bip);
                                        targetElement = element;
                                        if (param == null && typeElem != null)
                                        {
                                            param = typeElem.get_Parameter(bip);
                                            targetElement = typeElem;
                                        }
                                    }
                                }

                                if (param != null && !param.IsReadOnly)
                                {
                                    string currentValue = Utils.GetParameterValue(targetElement, paramName);
                                    if (currentValue != paramValue)
                                    {
                                        if (Utils.SetParameterValue(targetElement, paramName, paramValue))
                                        {
                                            elementUpdated = true;
                                            if (isType)
                                            {
                                                updatedTypeParameters.Add(Tuple.Create(typeElem.Id, paramName));
                                            }
                                        }
                                        else
                                        {
                                            ErrorMessages.Add(new ImportErrorItem { ElementId = idString, Description = $"Error updating parameter '{paramName}' with value '{paramValue}'" });
                                        }
                                    }
                                }
                            }

                            if (elementUpdated)
                            {
                                UpdatedElementsCount++;
                            }
                        }
                        else
                        {
                            ErrorMessages.Add(new ImportErrorItem { ElementId = idString, Description = "Element ID not found in model" });
                        }

                        processedRows++;
                        int percentage = (int)((double)processedRows / totalRows * 100);
                        _form.Dispatcher.Invoke(() => _form.UpdateProgressBar(percentage));
                    }

                    t.Commit();
                }

                _form.Dispatcher.Invoke(() => _form.HandleImportCompletion());
            }
            catch (Exception ex)
            {
                _form.Dispatcher.Invoke(() =>
                {
                    _form.HideProgressBar();
                    TaskDialog.Show("Error", $"Failed to import parameters:\n{ex.Message}");
                });
            }
            finally
            {
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