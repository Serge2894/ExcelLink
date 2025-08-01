using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelLink.Forms;
using ExcelLink.Common;
using ExcelLink.Properties;
using FilterTreeControlWPF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

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

                // Check if user clicked Export or Import
                if (form.btnExport.IsDefault) // Export was clicked
                {
                    return ExportToExcel(doc, selectedCategories, selectedParameters, isEntireModel);
                }
                else // Import was clicked
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
            saveDialog.Title = "Save Excel File";
            saveDialog.FileName = "RevitParameterExport.xlsx";

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

                int sheetIndex = 1;

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

                    // Write headers
                    worksheet.Cells[1, 1] = "Element ID";
                    for (int i = 0; i < selectedParameters.Count; i++)
                    {
                        worksheet.Cells[1, i + 2] = selectedParameters[i];
                    }

                    // Format headers
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, selectedParameters.Count + 1]];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    // Write element data
                    int row = 2;
                    foreach (Element element in elements)
                    {
                        // Write Element ID
                        worksheet.Cells[row, 1] = element.Id.IntegerValue.ToString();

                        // Write parameter values
                        for (int col = 0; col < selectedParameters.Count; col++)
                        {
                            string paramName = selectedParameters[col];
                            string value = GetParameterValue(element, paramName);
                            worksheet.Cells[row, col + 2] = value;
                        }

                        row++;
                    }

                    // Auto-fit columns
                    worksheet.Columns.AutoFit();

                    sheetIndex++;
                }

                // Save and close Excel
                workbook.SaveAs(excelFile);
                workbook.Close();
                excel.Quit();

                // Release COM objects
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                TaskDialog.Show("Success", $"Parameters exported successfully to:\n{excelFile}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Clean up Excel if error occurs
                try
                {
                    workbook.Close(false);
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