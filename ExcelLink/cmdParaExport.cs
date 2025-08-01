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

namespace ExcelLink
{
    [Transaction(TransactionMode.Manual)]
    public class cmdParaExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Show the form
                frmParaExport form = new frmParaExport(doc);
                bool? dialogResult = form.ShowDialog();

                if (dialogResult != true)
                {
                    return Result.Cancelled;
                }

                // Get selected categories and parameters
                List<string> selectedCategories = form.SelectedCategoryNames;
                List<string> selectedParameters = form.SelectedParameterNames;
                bool isEntireModel = form.IsEntireModelChecked;

                // Get the button that was clicked to trigger the dialog result
                bool isExportClicked = (bool)form.btnExport.Tag;

                // Check if user clicked Export or Import
                if (isExportClicked)
                {
                    return ExportToExcel(doc, selectedCategories, selectedParameters, isEntireModel);
                }
                else
                {
                    return ImportFromExcel(doc, selectedCategories, selectedParameters, isEntireModel);
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Result ExportToExcel(Document doc, List<string> selectedCategories, List<string> selectedParameters, bool isEntireModel)
        {
            // Validate selections
            if (!selectedCategories.Any())
            {
                TaskDialog.Show("Error", "Please select at least one category.");
                return Result.Failed;
            }

            if (!selectedParameters.Any())
            {
                TaskDialog.Show("Error", "Please select at least one parameter.");
                return Result.Failed;
            }

            // Prompt user to save Excel file
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel files|*.xlsx";
            saveDialog.Title = "Save Revit Parameters to Excel";

            // Set the default file name to the Revit project title
            string defaultFileName = doc.Title;
            if (string.IsNullOrEmpty(defaultFileName))
            {
                defaultFileName = "RevitParameterExport";
            }
            saveDialog.FileName = defaultFileName + ".xlsx";

            if (saveDialog.ShowDialog() != DialogResult.OK)
            {
                return Result.Cancelled;
            }

            string excelFile = saveDialog.FileName;

            // Create Excel application
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add();

            try
            {
                // Remove default sheets except the first one
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
                }

                // Create the Color Legend Sheet first
                Excel.Worksheet colorLegendSheet = (Excel.Worksheet)workbook.Worksheets[1];
                colorLegendSheet.Name = "Color Legend";

                // Write legend headers
                ((Excel.Range)colorLegendSheet.Cells[1, 2]).Value2 = "Color Legend";
                ((Excel.Range)colorLegendSheet.Cells[3, 2]).Value2 = "Color";
                ((Excel.Range)colorLegendSheet.Cells[3, 3]).Value2 = "Description";
                ((Excel.Range)colorLegendSheet.Cells[3, 4]).Value2 = "Notes";

                // Write legend content
                Excel.Range lightGrayCell = (Excel.Range)colorLegendSheet.Cells[4, 2];
                lightGrayCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                lightGrayCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)colorLegendSheet.Cells[4, 3]).Value2 = "Parameter does not exist for this element";
                ((Excel.Range)colorLegendSheet.Cells[4, 4]).Value2 = "Do not fill or edit cell";

                Excel.Range paleGoldenrodCell = (Excel.Range)colorLegendSheet.Cells[5, 2];
                paleGoldenrodCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.PaleGoldenrod);
                paleGoldenrodCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)colorLegendSheet.Cells[5, 3]).Value2 = "Type value";
                ((Excel.Range)colorLegendSheet.Cells[5, 4]).Value2 = "Type parameters with the same ID should be filled the same";

                Excel.Range yellowCell = (Excel.Range)colorLegendSheet.Cells[6, 2];
                yellowCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                yellowCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)colorLegendSheet.Cells[6, 3]).Value2 = "Read-only value";
                ((Excel.Range)colorLegendSheet.Cells[6, 4]).Value2 = "Uneditable cell";

                // Format legend headers
                Excel.Range legendHeaderRange = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[3, 4]];
                legendHeaderRange.Font.Bold = true;

                colorLegendSheet.Columns.AutoFit();

                int sheetIndex = 2; // Start with the second sheet for categories

                // Process each category
                foreach (string categoryName in selectedCategories)
                {
                    // Get category
                    Category category = GetCategoryByName(doc, categoryName);
                    if (category == null) continue;

                    // Get elements in category
                    FilteredElementCollector collector;
                    if (isEntireModel)
                    {
                        collector = new FilteredElementCollector(doc);
                    }
                    else
                    {
                        collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                    }

                    collector.OfCategoryId(category.Id);
                    collector.WhereElementIsNotElementType();

                    List<Element> elements = collector.ToList();

                    if (!elements.Any()) continue;

                    // Create or get worksheet
                    Excel.Worksheet worksheet;
                    if (sheetIndex == 1)
                    {
                        worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                    }
                    else
                    {
                        worksheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                    }

                    // Set sheet name (Excel limits sheet names to 31 characters)
                    string sheetName = categoryName.Length > 31 ? categoryName.Substring(0, 31) : categoryName;
                    worksheet.Name = sheetName;

                    // Write headers with multi-line text
                    ((Excel.Range)worksheet.Cells[1, 1]).Value2 = "Element ID";
                    for (int i = 0; i < selectedParameters.Count; i++)
                    {
                        string paramName = selectedParameters[i];
                        string paramType = "N/A";
                        string paramStorageType = "N/A";

                        Parameter param = elements.First().LookupParameter(paramName);
                        if (param == null)
                        {
                            Element typeElem = doc.GetElement(elements.First().GetTypeId());
                            param = typeElem?.LookupParameter(paramName);
                            if (param != null)
                            {
                                paramType = "Type Parameter";
                            }
                            else
                            {
                                param = elements.First().get_Parameter(GetBuiltInParameterByName(paramName));
                                if (param != null)
                                {
                                    paramType = "Instance Parameter";
                                }
                            }
                        }
                        else
                        {
                            paramType = "Instance Parameter";
                        }

                        if (param != null)
                        {
                            paramStorageType = GetParameterStorageTypeString(param.StorageType);
                        }

                        string headerText = $"{paramName}{Environment.NewLine}({paramType}){Environment.NewLine}Type: {paramStorageType}";
                        ((Excel.Range)worksheet.Cells[1, i + 2]).Value2 = headerText;
                    }

                    // Format headers
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, selectedParameters.Count + 1]];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    headerRange.WrapText = true;
                    headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    ((Excel.Range)worksheet.Rows[1]).AutoFit();


                    // Write element data
                    int row = 2;
                    foreach (Element element in elements)
                    {
                        // Write Element ID and color it yellow for Read-only
                        Excel.Range idCell = (Excel.Range)worksheet.Cells[row, 1];
                        idCell.Value2 = element.Id.IntegerValue.ToString();
                        idCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                        // Write parameter values
                        for (int col = 0; col < selectedParameters.Count; col++)
                        {
                            string paramName = selectedParameters[col];
                            Excel.Range dataCell = (Excel.Range)worksheet.Cells[row, col + 2];

                            Parameter param = element.LookupParameter(paramName);
                            string value = string.Empty;
                            bool isTypeParam = false;

                            // Check if the parameter exists as an instance parameter
                            if (param != null)
                            {
                                value = GetParameterValue(element, paramName);
                            }
                            else
                            {
                                // Check if the parameter exists as a type parameter
                                Element typeElem = doc.GetElement(element.GetTypeId());
                                if (typeElem != null)
                                {
                                    param = typeElem.LookupParameter(paramName);
                                    if (param != null)
                                    {
                                        value = GetParameterValue(typeElem, paramName);
                                        isTypeParam = true;
                                    }
                                }
                            }

                            if (param != null)
                            {
                                dataCell.Value2 = value;
                                if (isTypeParam)
                                {
                                    dataCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.PaleGoldenrod);
                                }
                            }
                            else
                            {
                                // If parameter does not exist, color the cell grey
                                dataCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            }
                        }

                        row++;
                    }

                    // Auto-fit columns
                    worksheet.Columns.AutoFit();

                    sheetIndex++;
                }

                // Save the file
                workbook.SaveAs(excelFile);

                // Activate the Color Legend sheet
                ((Excel.Worksheet)workbook.Worksheets["Color Legend"]).Activate();

                // Open the file and keep Excel visible
                excel.Visible = true;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Clean up Excel if error occurs
                try
                {
                    if (workbook != null) workbook.Close(false);
                    excel.Quit();
                }
                catch { }

                TaskDialog.Show("Error", $"Failed to export parameters:\n{ex.Message}");
                return Result.Failed;
            }
        }

        private Result ImportFromExcel(Document doc, List<string> selectedCategories, List<string> selectedParameters, bool isEntireModel)
        {
            // Prompt user to select Excel file
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel files|*.xlsx;*.xls";
            openDialog.Title = "Select Excel File to Import";

            if (openDialog.ShowDialog() != DialogResult.OK)
            {
                return Result.Cancelled;
            }

            string excelFile = openDialog.FileName;

            // Open Excel file
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                workbook = excel.Workbooks.Open(excelFile);

                using (Transaction trans = new Transaction(doc, "Import Parameters from Excel"))
                {
                    trans.Start();

                    int totalUpdated = 0;

                    // Process each worksheet
                    foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                        string categoryName = worksheet.Name;

                        // Check if this category was selected
                        if (!selectedCategories.Contains(categoryName))
                        {
                            continue;
                        }

                        // Get used range
                        Excel.Range usedRange = worksheet.UsedRange;
                        int rowCount = usedRange.Rows.Count;
                        int colCount = usedRange.Columns.Count;

                        if (rowCount < 2) continue; // Skip if no data rows

                        // Read headers
                        List<string> headers = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            object headerValue = ((Excel.Range)worksheet.Cells[1, col]).Value2;
                            headers.Add(headerValue?.ToString() ?? "");
                        }

                        // Process data rows
                        for (int row = 2; row <= rowCount; row++)
                        {
                            // Get Element ID from first column
                            object idValue = ((Excel.Range)worksheet.Cells[row, 1]).Value2;
                            if (idValue == null) continue;

                            string elementIdStr = idValue.ToString();
                            if (!int.TryParse(elementIdStr, out int elementIdInt)) continue;

                            ElementId elementId = new ElementId(elementIdInt);
                            Element element = doc.GetElement(elementId);

                            if (element == null) continue;

                            // Update parameters
                            for (int col = 2; col <= colCount; col++)
                            {
                                string paramName = headers[col - 1];
                                if (!selectedParameters.Contains(paramName)) continue;

                                object cellValue = ((Excel.Range)worksheet.Cells[row, col]).Value2;
                                if (cellValue != null)
                                {
                                    bool updated = SetParameterValue(element, paramName, cellValue.ToString());
                                    if (updated) totalUpdated++;
                                }
                            }
                        }
                    }

                    trans.Commit();

                    TaskDialog.Show("Success", $"Import completed successfully.\n{totalUpdated} parameter values updated.");
                }

                // Close Excel
                workbook.Close(false);
                excel.Quit();

                // Release COM objects
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Clean up Excel if error occurs
                try
                {
                    if (workbook != null) workbook.Close(false);
                    excel.Quit();
                }
                catch { }

                TaskDialog.Show("Error", $"Failed to import parameters:\n{ex.Message}");
                return Result.Failed;
            }
        }

        private string GetParameterStorageTypeString(StorageType storageType)
        {
            switch (storageType)
            {
                case StorageType.Integer:
                    return "Integer";
                case StorageType.Double:
                    return "Decimal";
                case StorageType.String:
                    return "Text";
                case StorageType.ElementId:
                    return "Element ID";
                default:
                    return "Other";
            }
        }

        private BuiltInParameter GetBuiltInParameterByName(string paramName)
        {
            if (Enum.TryParse(paramName, out BuiltInParameter bip))
            {
                return bip;
            }
            return BuiltInParameter.INVALID;
        }

        private Category GetCategoryByName(Document doc, string categoryName)
        {
            Categories categories = doc.Settings.Categories;
            foreach (Category category in categories)
            {
                if (category.Name == categoryName)
                {
                    return category;
                }
            }
            return null;
        }

        private string GetParameterValue(Element element, string parameterName)
        {
            // Try instance parameter first
            Parameter param = element.LookupParameter(parameterName);

            // If not found, try type parameter
            if (param == null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element elementType = element.Document.GetElement(typeId);
                    if (elementType != null)
                    {
                        param = elementType.LookupParameter(parameterName);
                    }
                }
            }

            if (param == null) return "";

            // Get value based on storage type
            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? "";
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return param.AsValueString() ?? param.AsDouble().ToString();
                case StorageType.ElementId:
                    ElementId id = param.AsElementId();
                    if (id != ElementId.InvalidElementId)
                    {
                        Element elem = element.Document.GetElement(id);
                        return elem?.Name ?? id.IntegerValue.ToString();
                    }
                    return "";
                default:
                    return "";
            }
        }

        private bool SetParameterValue(Element element, string parameterName, string value)
        {
            // Try instance parameter first
            Parameter param = element.LookupParameter(parameterName);
            bool isTypeParam = false;

            // If not found, try type parameter
            if (param == null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element elementType = element.Document.GetElement(typeId);
                    if (elementType != null)
                    {
                        param = elementType.LookupParameter(parameterName);
                        isTypeParam = true;
                    }
                }
            }

            if (param == null || param.IsReadOnly) return false;

            try
            {
                // Set value based on storage type
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        return true;
                    case StorageType.Integer:
                        if (int.TryParse(value, out int intValue))
                        {
                            param.Set(intValue);
                            return true;
                        }
                        break;
                    case StorageType.Double:
                        if (double.TryParse(value, out double doubleValue))
                        {
                            param.Set(doubleValue);
                            return true;
                        }
                        break;
                    case StorageType.ElementId:
                        // For ElementId parameters, try to parse as integer
                        if (int.TryParse(value, out int idValue))
                        {
                            param.Set(new ElementId(idValue));
                            return true;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                // Log error but continue processing
                System.Diagnostics.Debug.WriteLine($"Failed to set parameter {parameterName}: {ex.Message}");
            }

            return false;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
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