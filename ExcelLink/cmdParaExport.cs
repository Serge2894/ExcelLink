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

                // Merge and center title
                Excel.Range titleRange = colorLegendSheet.Range[colorLegendSheet.Cells[1, 2], colorLegendSheet.Cells[1, 4]];
                titleRange.Merge();
                titleRange.Value2 = "Color Legend";
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;

                // Add thick border around title
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                titleRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                // Write legend headers
                ((Excel.Range)colorLegendSheet.Cells[3, 2]).Value2 = "Color";
                ((Excel.Range)colorLegendSheet.Cells[3, 3]).Value2 = "Description";
                ((Excel.Range)colorLegendSheet.Cells[3, 4]).Value2 = "Notes";

                // Format headers
                Excel.Range legendHeaderRange = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[3, 4]];
                legendHeaderRange.Font.Bold = true;
                legendHeaderRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                legendHeaderRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // Write legend content - Row 4: Grey (#D3D3D3)
                Excel.Range greyCell = (Excel.Range)colorLegendSheet.Cells[4, 2];
                greyCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                ((Excel.Range)colorLegendSheet.Cells[4, 3]).Value2 = "Parameter does not exist for this element";
                ((Excel.Range)colorLegendSheet.Cells[4, 4]).Value2 = "Do not fill or edit cell";

                // Row 5: Light Yellow (#FFE699)
                Excel.Range lightYellowCell = (Excel.Range)colorLegendSheet.Cells[5, 2];
                lightYellowCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                ((Excel.Range)colorLegendSheet.Cells[5, 3]).Value2 = "Type value";
                ((Excel.Range)colorLegendSheet.Cells[5, 4]).Value2 = "Type parameters with the same ID should be filled the same";

                // Row 6: Red (#FF4747)
                Excel.Range redCell = (Excel.Range)colorLegendSheet.Cells[6, 2];
                redCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                ((Excel.Range)colorLegendSheet.Cells[6, 3]).Value2 = "Read-only value";
                ((Excel.Range)colorLegendSheet.Cells[6, 4]).Value2 = "Uneditable cell";

                // Apply borders to all data cells
                Excel.Range dataRange = colorLegendSheet.Range[colorLegendSheet.Cells[4, 2], colorLegendSheet.Cells[6, 4]];
                dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                // Apply thick outside border to the entire table
                Excel.Range entireTable = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[6, 4]];
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                // Set column widths
                ((Excel.Range)colorLegendSheet.Columns[2]).ColumnWidth = 15;
                ((Excel.Range)colorLegendSheet.Columns[3]).ColumnWidth = 40;
                ((Excel.Range)colorLegendSheet.Columns[4]).ColumnWidth = 50;

                // Center align the color column
                Excel.Range colorColumn = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[6, 2]];
                colorColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

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
                    Excel.Range elementIdHeader = (Excel.Range)worksheet.Cells[1, 1];
                    elementIdHeader.Value2 = "Element ID";
                    elementIdHeader.ColumnWidth = 12;

                    for (int i = 0; i < selectedParameters.Count; i++)
                    {
                        string paramName = selectedParameters[i];
                        string paramType = "N/A";
                        string paramStorageType = "N/A";

                        // First check instance parameter
                        Parameter param = elements.First().LookupParameter(paramName);
                        bool isTypeParam = false;

                        if (param != null)
                        {
                            paramType = "Instance Parameter";
                        }
                        else
                        {
                            // Check type parameter
                            Element typeElem = doc.GetElement(elements.First().GetTypeId());
                            if (typeElem != null)
                            {
                                param = typeElem.LookupParameter(paramName);
                                if (param != null)
                                {
                                    paramType = "Type Parameter";
                                    isTypeParam = true;
                                }
                            }
                        }

                        // If still not found, check for built-in parameters
                        if (param == null)
                        {
                            // Check built-in instance parameters
                            BuiltInParameter bip = GetBuiltInParameterByName(paramName);
                            if (bip != BuiltInParameter.INVALID)
                            {
                                param = elements.First().get_Parameter(bip);
                                if (param != null)
                                {
                                    paramType = "Instance Parameter";
                                }
                                else
                                {
                                    // Check if it's a built-in type parameter
                                    Element typeElem = doc.GetElement(elements.First().GetTypeId());
                                    if (typeElem != null)
                                    {
                                        param = typeElem.get_Parameter(bip);
                                        if (param != null)
                                        {
                                            paramType = "Type Parameter";
                                            isTypeParam = true;
                                        }
                                    }
                                }
                            }
                        }

                        if (param != null)
                        {
                            paramStorageType = GetParameterStorageTypeString(param.StorageType);
                        }

                        string headerText = $"{paramName}{Environment.NewLine}({paramType}){Environment.NewLine}Type: {paramStorageType}";
                        Excel.Range headerCell = (Excel.Range)worksheet.Cells[1, i + 2];
                        headerCell.Value2 = headerText;

                        // Set column width based on parameter name length, but with reasonable limits
                        int columnWidth = Math.Max(15, Math.Min(30, paramName.Length + 5));
                        headerCell.ColumnWidth = columnWidth;
                    }

                    // Format headers
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, selectedParameters.Count + 1]];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
                    headerRange.WrapText = true;
                    headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    headerRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    // Set row height to accommodate 3 lines of text
                    ((Excel.Range)worksheet.Rows[1]).RowHeight = 45;


                    // Write element data
                    int row = 2;
                    foreach (Element element in elements)
                    {
                        // Write Element ID and color it grey (#D3D3D3) for Read-only
                        Excel.Range idCell = (Excel.Range)worksheet.Cells[row, 1];
                        idCell.Value2 = element.Id.IntegerValue.ToString();
                        idCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                        idCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        idCell.Borders.Weight = Excel.XlBorderWeight.xlThin;

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

                            // If still not found, check built-in parameters
                            if (param == null)
                            {
                                BuiltInParameter bip = GetBuiltInParameterByName(paramName);
                                if (bip != BuiltInParameter.INVALID)
                                {
                                    param = element.get_Parameter(bip);
                                    if (param != null)
                                    {
                                        value = GetParameterValue(element, paramName);
                                    }
                                    else
                                    {
                                        // Check if it's a built-in type parameter
                                        Element typeElem = doc.GetElement(element.GetTypeId());
                                        if (typeElem != null)
                                        {
                                            param = typeElem.get_Parameter(bip);
                                            if (param != null)
                                            {
                                                value = GetParameterValue(typeElem, paramName);
                                                isTypeParam = true;
                                            }
                                        }
                                    }
                                }
                            }

                            if (param != null)
                            {
                                dataCell.Value2 = value;
                                // Special handling for Family and Family and Type - they should always be red
                                if (paramName == "Family" || paramName == "Family and Type")
                                {
                                    // These are always read-only, color them red
                                    dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                                }
                                else if (param.IsReadOnly)
                                {
                                    // Other read-only parameters get red color
                                    dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                                }
                                else if (isTypeParam)
                                {
                                    // Type parameters get light yellow color
                                    dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                                }
                                // If the parameter is editable but no color assigned, leave it white (no color)
                                // Note: Workset is an editable instance parameter, so it will be white
                            }
                            else
                            {
                                // If parameter does not exist, color the cell grey (#D3D3D3)
                                dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                            }

                            // Add borders to all data cells
                            dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        row++;
                    }

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

                // Dictionary to store errors
                Dictionary<string, List<string>> importErrors = new Dictionary<string, List<string>>();

                using (Transaction trans = new Transaction(doc, "Import Parameters from Excel"))
                {
                    trans.Start();

                    int totalUpdated = 0;
                    int totalSkipped = 0;

                    // Process each worksheet
                    foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                        string categoryName = worksheet.Name;

                        // Skip Color Legend sheet
                        if (categoryName == "Color Legend")
                            continue;

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

                        // Read headers - extract parameter names from multi-line headers
                        List<string> parameterNames = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            object headerValue = ((Excel.Range)worksheet.Cells[1, col]).Value2;
                            string headerText = headerValue?.ToString() ?? "";

                            if (col == 1)
                            {
                                parameterNames.Add("Element ID");
                            }
                            else
                            {
                                // Extract parameter name from multi-line header
                                string[] lines = headerText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                if (lines.Length > 0)
                                {
                                    parameterNames.Add(lines[0]); // First line is the parameter name
                                }
                                else
                                {
                                    parameterNames.Add("");
                                }
                            }
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
                                string paramName = parameterNames[col - 1];
                                if (string.IsNullOrEmpty(paramName) || !selectedParameters.Contains(paramName))
                                    continue;

                                // Get cell and its color
                                Excel.Range cell = (Excel.Range)worksheet.Cells[row, col];
                                object cellValue = cell.Value2;
                                double cellColor = cell.Interior.Color;

                                // Convert Excel color to RGB
                                int colorInt = Convert.ToInt32(cellColor);
                                System.Drawing.Color color = System.Drawing.ColorTranslator.FromOle(colorInt);
                                string htmlColor = System.Drawing.ColorTranslator.ToHtml(color).ToUpper();

                                // Skip if cell is grey (parameter doesn't exist) or red (read-only)
                                if (htmlColor == "#D3D3D3" || htmlColor == "#FF4747")
                                {
                                    totalSkipped++;
                                    continue;
                                }

                                if (cellValue != null && cellValue.ToString().Trim() != "")
                                {
                                    UpdateResult result = SetParameterValueWithValidation(element, paramName, cellValue.ToString());

                                    if (result.Success)
                                    {
                                        totalUpdated++;
                                    }
                                    else
                                    {
                                        // Add to error list
                                        string errorKey = $"{categoryName} - {paramName}";
                                        if (!importErrors.ContainsKey(errorKey))
                                        {
                                            importErrors[errorKey] = new List<string>();
                                        }
                                        importErrors[errorKey].Add($"Element ID: {elementIdStr} - {result.ErrorMessage}");
                                    }
                                }
                            }
                        }
                    }

                    trans.Commit();

                    // Show results
                    StringBuilder resultMessage = new StringBuilder();
                    resultMessage.AppendLine($"Import completed successfully.");
                    resultMessage.AppendLine($"{totalUpdated} parameter values updated.");
                    resultMessage.AppendLine($"{totalSkipped} cells skipped (read-only or non-existent parameters).");

                    if (importErrors.Any())
                    {
                        resultMessage.AppendLine("\nThe following parameters could not be updated:");
                        foreach (var error in importErrors)
                        {
                            resultMessage.AppendLine($"\n{error.Key}:");
                            foreach (var elementError in error.Value)
                            {
                                resultMessage.AppendLine($"  - {elementError}");
                            }
                        }

                        // Show detailed error report
                        TaskDialogResult result = TaskDialog.Show("Import Results - Errors Found",
                            resultMessage.ToString(),
                            TaskDialogCommonButtons.Ok);
                    }
                    else
                    {
                        TaskDialog.Show("Import Results", resultMessage.ToString());
                    }
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

        private class UpdateResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
        }

        private UpdateResult SetParameterValueWithValidation(Element element, string parameterName, string value)
        {
            UpdateResult result = new UpdateResult { Success = false };

            // Try instance parameter first
            Parameter param = element.LookupParameter(parameterName);
            Element targetElement = element;

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
                        targetElement = elementType;
                    }
                }
            }

            // If still not found, try built-in parameters
            if (param == null)
            {
                BuiltInParameter bip = GetBuiltInParameterByName(parameterName);
                if (bip != BuiltInParameter.INVALID)
                {
                    param = element.get_Parameter(bip);
                    targetElement = element;

                    // If not found on instance, try on type
                    if (param == null)
                    {
                        ElementId typeId = element.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element elementType = element.Document.GetElement(typeId);
                            if (elementType != null)
                            {
                                param = elementType.get_Parameter(bip);
                                targetElement = elementType;
                            }
                        }
                    }
                }
            }

            if (param == null)
            {
                result.ErrorMessage = "Parameter not found";
                return result;
            }

            if (param.IsReadOnly)
            {
                result.ErrorMessage = "Parameter is read-only";
                return result;
            }

            try
            {
                // Set value based on storage type
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        result.Success = true;
                        break;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intValue))
                        {
                            param.Set(intValue);
                            result.Success = true;
                        }
                        else
                        {
                            result.ErrorMessage = $"Invalid integer value: '{value}'";
                        }
                        break;

                    case StorageType.Double:
                        if (double.TryParse(value, out double doubleValue))
                        {
                            param.Set(doubleValue);
                            result.Success = true;
                        }
                        else
                        {
                            result.ErrorMessage = $"Invalid decimal value: '{value}'";
                        }
                        break;

                    case StorageType.ElementId:
                        // Special handling for Workset parameter
                        if (paramName == "Workset")
                        {
                            // Find workset by name
                            FilteredWorksetCollector worksetCollector = new FilteredWorksetCollector(element.Document);
                            Workset workset = worksetCollector.FirstOrDefault(w => w.Name == value);

                            if (workset != null)
                            {
                                param.Set(workset.Id);
                                result.Success = true;
                            }
                            else
                            {
                                result.ErrorMessage = $"Workset '{value}' not found";
                            }
                        }
                        else
                        {
                            // For other ElementId parameters, try to parse as integer
                            if (int.TryParse(value, out int idValue))
                            {
                                param.Set(new ElementId(idValue));
                                result.Success = true;
                            }
                            else
                            {
                                result.ErrorMessage = $"Invalid Element ID value: '{value}'";
                            }
                        }
                        break;

                    default:
                        result.ErrorMessage = $"Unsupported parameter type: {param.StorageType}";
                        break;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Error setting value: {ex.Message}";
            }

            return result;
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
            // Direct enum parsing
            if (Enum.TryParse(paramName, out BuiltInParameter bip))
            {
                return bip;
            }

            // Common parameter name mappings
            switch (paramName)
            {
                // Basic parameters
                case "Family":
                    return BuiltInParameter.ELEM_FAMILY_PARAM;
                case "Family and Type":
                    return BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM;
                case "Type":
                    return BuiltInParameter.ELEM_TYPE_PARAM;
                case "Type Id":
                    return BuiltInParameter.SYMBOL_ID_PARAM;

                // Workset parameter
                case "Workset":
                    return BuiltInParameter.ELEM_PARTITION_PARAM;

                // Comments and descriptions
                case "Type Comments":
                    return BuiltInParameter.ALL_MODEL_TYPE_COMMENTS;
                case "Comments":
                    return BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS;
                case "Mark":
                    return BuiltInParameter.ALL_MODEL_MARK;
                case "Type Mark":
                    return BuiltInParameter.ALL_MODEL_TYPE_MARK;
                case "Description":
                    return BuiltInParameter.ALL_MODEL_DESCRIPTION;

                // URL and manufacturer info
                case "URL":
                    return BuiltInParameter.ALL_MODEL_URL;
                case "Type Name":
                    return BuiltInParameter.SYMBOL_NAME_PARAM;
                case "Manufacturer":
                    return BuiltInParameter.ALL_MODEL_MANUFACTURER;
                case "Model":
                    return BuiltInParameter.ALL_MODEL_MODEL;
                case "Cost":
                    return BuiltInParameter.ALL_MODEL_COST;

                // Images
                case "Image":
                    return BuiltInParameter.ALL_MODEL_IMAGE;
                case "Type Image":
                    return BuiltInParameter.ALL_MODEL_TYPE_IMAGE;

                // Assembly and classification
                case "Assembly Code":
                    return BuiltInParameter.UNIFORMAT_CODE;
                case "Assembly Description":
                    return BuiltInParameter.UNIFORMAT_DESCRIPTION;
                case "Keynote":
                    return BuiltInParameter.KEYNOTE_PARAM;
                case "OmniClass Number":
                    return BuiltInParameter.OMNICLASS_CODE;
                case "OmniClass Title":
                    return BuiltInParameter.OMNICLASS_DESCRIPTION;

                // Room specific parameters
                case "Name":
                    return BuiltInParameter.ROOM_NAME;
                case "Number":
                    return BuiltInParameter.ROOM_NUMBER;
                case "Department":
                    return BuiltInParameter.ROOM_DEPARTMENT;
                case "Occupancy":
                    return BuiltInParameter.ROOM_OCCUPANCY;
                case "Occupant":
                    return BuiltInParameter.ROOM_OCCUPANT;
                case "Base Finish":
                    return BuiltInParameter.ROOM_FINISH_BASE;
                case "Ceiling Finish":
                    return BuiltInParameter.ROOM_FINISH_CEILING;
                case "Wall Finish":
                    return BuiltInParameter.ROOM_FINISH_WALL;
                case "Floor Finish":
                    return BuiltInParameter.ROOM_FINISH_FLOOR;

                // Floor specific parameters
                case "Default Thickness":
                    return BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM;
                case "Thickness":
                    return BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM;
                case "Function":
                    return BuiltInParameter.FUNCTION_PARAM;
                case "Structural":
                    return BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
                case "Structural Usage":
                    return BuiltInParameter.FLOOR_PARAM_STRUCTURAL_USAGE;

                // Wall specific parameters
                case "Width":
                    return BuiltInParameter.WALL_ATTR_WIDTH_PARAM;
                case "Wall Structural Usage":
                    return BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM;

                // General parameters
                case "Area":
                    return BuiltInParameter.HOST_AREA_COMPUTED;
                case "Volume":
                    return BuiltInParameter.HOST_VOLUME_COMPUTED;
                case "Perimeter":
                    return BuiltInParameter.HOST_PERIMETER_COMPUTED;
                case "Length":
                    return BuiltInParameter.CURVE_ELEM_LENGTH;
                case "Level":
                    return BuiltInParameter.LEVEL_PARAM;
                case "Base Level":
                    return BuiltInParameter.LEVEL_PARAM;
                case "Top Level":
                    return BuiltInParameter.WALL_HEIGHT_TYPE;
                case "Base Offset":
                    return BuiltInParameter.WALL_BASE_OFFSET;
                case "Top Offset":
                    return BuiltInParameter.WALL_TOP_OFFSET;
                case "Height":
                    return BuiltInParameter.WALL_USER_HEIGHT_PARAM;

                // Door/Window parameters
                case "Code Name":
                    return BuiltInParameter.DOOR_NUMBER;
                case "Head Height":
                    return BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM;
                case "Sill Height":
                    return BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM;

                default:
                    return BuiltInParameter.INVALID;
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

            // If still not found, try built-in parameters
            if (param == null)
            {
                BuiltInParameter bip = GetBuiltInParameterByName(parameterName);
                if (bip != BuiltInParameter.INVALID)
                {
                    param = element.get_Parameter(bip);

                    // If not found on instance, try on type
                    if (param == null)
                    {
                        ElementId typeId = element.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element elementType = element.Document.GetElement(typeId);
                            if (elementType != null)
                            {
                                param = elementType.get_Parameter(bip);
                            }
                        }
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
                        // Special handling for Workset parameter
                        if (parameterName == "Workset")
                        {
                            Workset workset = element.Document.GetWorksetTable().GetWorkset(id);
                            return workset?.Name ?? id.IntegerValue.ToString();
                        }
                        else
                        {
                            Element elem = element.Document.GetElement(id);
                            return elem?.Name ?? id.IntegerValue.ToString();
                        }
                    }
                    return "";
                default:
                    return "";
            }
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