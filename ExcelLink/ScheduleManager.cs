using Autodesk.Revit.DB;
using ExcelLink.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLink.Common
{
    /// <summary>
    /// Stores properties for each column in a schedule.
    /// </summary>
    public class ColumnProperty
    {
        public bool IsReadOnly { get; set; }
        public bool IsType { get; set; }
        public bool IsCalculated { get; set; } // Added for calculated fields
        public string FieldType { get; set; } // Added to identify field type
    }

    /// <summary>
    /// A simple data class to hold schedule information, free of any Revit API objects.
    /// </summary>
    public class SimpleScheduleData
    {
        public string Name { get; set; }
        public List<string> Headers { get; set; }
        public List<string> ColumnLetters { get; set; }
        public List<ColumnProperty> ColumnProperties { get; set; }
        public List<List<string>> BodyRows { get; set; }
        public List<List<string>> SummaryRows { get; set; }
        public List<bool> IsGroupHeaderOrFooterRow { get; set; }
        public List<bool> IsBlankLineRow { get; set; }
        // Added to store field information
        public List<ScheduleFieldInfo> FieldInfos { get; set; }
    }

    /// <summary>
    /// Information about a schedule field
    /// </summary>
    public class ScheduleFieldInfo
    {
        public string Name { get; set; }
        public bool IsCalculatedField { get; set; }
        public bool IsSharedParameter { get; set; }
        public bool IsCount { get; set; }
        public ScheduleFieldType FieldType { get; set; }
        public ElementId ParameterId { get; set; }
    }

    public class ScheduleManager
    {
        private Document _doc;

        public ScheduleManager(Document doc)
        {
            _doc = doc;
        }

        public List<ViewSchedule> GetAllSchedules()
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc);
            var schedules = collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTemplate && !s.IsTitleblockRevisionSchedule && s.Definition != null)
                .OrderBy(s => s.Name)
                .ToList();

            return schedules;
        }

        public List<SimpleScheduleData> GetScheduleDataForExport(List<ViewSchedule> schedules, bool includeHeaders, bool includeGrandTotals)
        {
            var allScheduleData = new List<SimpleScheduleData>();

            foreach (var schedule in schedules)
            {
                var simpleData = new SimpleScheduleData
                {
                    Name = schedule.Name,
                    Headers = new List<string>(),
                    ColumnLetters = new List<string>(),
                    ColumnProperties = new List<ColumnProperty>(),
                    BodyRows = new List<List<string>>(),
                    SummaryRows = new List<List<string>>(),
                    IsGroupHeaderOrFooterRow = new List<bool>(),
                    IsBlankLineRow = new List<bool>(),
                    FieldInfos = new List<ScheduleFieldInfo>()
                };

                ScheduleDefinition definition = schedule.Definition;
                TableData tableData = schedule.GetTableData();
                TableSectionData bodySection = tableData.GetSectionData(SectionType.Body);

                var visibleFields = new List<ScheduleField>();
                for (int i = 0; i < definition.GetFieldCount(); i++)
                {
                    var field = definition.GetField(i);
                    if (!field.IsHidden)
                    {
                        visibleFields.Add(field);
                    }
                }

                int numberOfRows = bodySection.NumberOfRows;
                int numberOfColumns = visibleFields.Count > 0 ? visibleFields.Count : bodySection.NumberOfColumns;

                if (numberOfColumns != bodySection.NumberOfColumns)
                {
                    numberOfColumns = bodySection.NumberOfColumns;
                }

                Element sampleElement = new FilteredElementCollector(_doc, schedule.Id).FirstElement();
                Element sampleTypeElement = sampleElement != null ? _doc.GetElement(sampleElement.GetTypeId()) : null;

                if (includeHeaders)
                {
                    for (int colIdx = 0; colIdx < numberOfColumns; colIdx++)
                    {
                        string headerText = schedule.GetCellText(SectionType.Header, 0, colIdx);
                        simpleData.Headers.Add(headerText);
                        simpleData.ColumnLetters.Add(GetExcelColumnName(colIdx + 1));

                        if (colIdx < visibleFields.Count)
                        {
                            var field = visibleFields[colIdx];
                            var fieldInfo = GetScheduleFieldInfo(field, definition);
                            simpleData.FieldInfos.Add(fieldInfo);

                            var colProp = new ColumnProperty
                            {
                                IsReadOnly = true,
                                IsType = false,
                                IsCalculated = fieldInfo.IsCalculatedField || fieldInfo.IsCount,
                                FieldType = fieldInfo.FieldType.ToString()
                            };

                            // For non-calculated fields, check if it's a parameter
                            if (!fieldInfo.IsCalculatedField && !fieldInfo.IsCount)
                            {
                                string paramName = field.GetName();

                                if (sampleElement != null)
                                {
                                    Parameter param = GetParameterByField(sampleElement, field);
                                    if (param != null)
                                    {
                                        colProp.IsReadOnly = param.IsReadOnly;
                                        colProp.IsType = false;
                                    }
                                    else if (sampleTypeElement != null)
                                    {
                                        param = GetParameterByField(sampleTypeElement, field);
                                        if (param != null)
                                        {
                                            colProp.IsReadOnly = param.IsReadOnly;
                                            colProp.IsType = true;
                                        }
                                    }
                                }
                            }
                            simpleData.ColumnProperties.Add(colProp);
                        }
                        else
                        {
                            simpleData.ColumnProperties.Add(new ColumnProperty
                            {
                                IsReadOnly = true,
                                IsType = false,
                                IsCalculated = false,
                                FieldType = "Unknown"
                            });
                            simpleData.FieldInfos.Add(new ScheduleFieldInfo
                            {
                                Name = headerText,
                                IsCalculatedField = false,
                                IsSharedParameter = false,
                                IsCount = false
                            });
                        }
                    }
                }

                // Export body rows
                for (int row = 0; row < numberOfRows; row++)
                {
                    var rowData = new List<string>();
                    for (int col = 0; col < numberOfColumns; col++)
                    {
                        rowData.Add(schedule.GetCellText(SectionType.Body, row, col));
                    }
                    simpleData.BodyRows.Add(rowData);
                }

                // Process row types
                foreach (var rowData in simpleData.BodyRows)
                {
                    bool isBlank = rowData.All(string.IsNullOrWhiteSpace);
                    simpleData.IsBlankLineRow.Add(isBlank);

                    if (isBlank)
                    {
                        simpleData.IsGroupHeaderOrFooterRow.Add(false);
                        continue;
                    }

                    int nonEmptyCount = rowData.Count(c => !string.IsNullOrWhiteSpace(c));
                    bool isHeaderOrFooter;

                    if (numberOfColumns > 2)
                    {
                        isHeaderOrFooter = (nonEmptyCount > 0 && nonEmptyCount <= 2);
                    }
                    else
                    {
                        isHeaderOrFooter = (nonEmptyCount == 1);
                    }
                    simpleData.IsGroupHeaderOrFooterRow.Add(isHeaderOrFooter);
                }

                if (includeGrandTotals)
                {
                    TableSectionData summarySection = tableData.GetSectionData(SectionType.Summary);
                    if (summarySection != null && summarySection.NumberOfRows > 0)
                    {
                        for (int row = 0; row < summarySection.NumberOfRows; row++)
                        {
                            var summaryRowData = new List<string>();
                            for (int col = 0; col < summarySection.NumberOfColumns; col++)
                            {
                                summaryRowData.Add(schedule.GetCellText(SectionType.Summary, row, col));
                            }
                            simpleData.SummaryRows.Add(summaryRowData);
                        }
                    }
                }
                allScheduleData.Add(simpleData);
            }
            return allScheduleData;
        }

        /// <summary>
        /// Get detailed information about a schedule field
        /// </summary>
        private ScheduleFieldInfo GetScheduleFieldInfo(ScheduleField field, ScheduleDefinition definition)
        {
            var fieldInfo = new ScheduleFieldInfo
            {
                Name = field.GetName(),
                IsCalculatedField = false,
                IsSharedParameter = false,
                IsCount = false,
                FieldType = field.FieldType,
                ParameterId = field.ParameterId
            };

            // Check if it's a Count field
            if (field.FieldType == ScheduleFieldType.Count)
            {
                fieldInfo.IsCount = true;
                fieldInfo.IsCalculatedField = true;
            }
            // Check if it's a calculated value field
            else if (field.IsCalculatedField)
            {
                fieldInfo.IsCalculatedField = true;
            }
            // Check if it's a shared parameter
            else if (field.ParameterId != null && field.ParameterId != ElementId.InvalidElementId)
            {
                try
                {
                    // Try to get the shared parameter definition
                    SharedParameterElement sharedParam = _doc.GetElement(field.ParameterId) as SharedParameterElement;
                    if (sharedParam != null)
                    {
                        fieldInfo.IsSharedParameter = true;
                    }
                }
                catch { }
            }

            return fieldInfo;
        }

        /// <summary>
        /// Get parameter from element based on schedule field
        /// </summary>
        private Parameter GetParameterByField(Element element, ScheduleField field)
        {
            if (element == null || field == null)
                return null;

            // First try to get by parameter ID if it exists
            if (field.ParameterId != null && field.ParameterId != ElementId.InvalidElementId)
            {
                // Check if it's a built-in parameter
                if (field.ParameterId.IntegerValue < 0)
                {
                    // It's a built-in parameter
                    BuiltInParameter builtInParam = (BuiltInParameter)field.ParameterId.IntegerValue;
                    Parameter param = element.get_Parameter(builtInParam);
                    if (param != null)
                        return param;
                }
                else
                {
                    // It's a shared or project parameter - use the GUID-based approach
                    foreach (Parameter param in element.Parameters)
                    {
                        if (param.Id == field.ParameterId)
                            return param;
                    }
                }
            }

            // Try to get by name
            string paramName = field.GetName();
            Parameter paramByName = element.LookupParameter(paramName);
            if (paramByName != null)
                return paramByName;

            // Try built-in parameters by name
            BuiltInParameter bipByName = Utils.GetBuiltInParameterByName(paramName);
            if (bipByName != BuiltInParameter.INVALID)
            {
                return element.get_Parameter(bipByName);
            }

            return null;
        }

        public void ExportSchedulesToExcel(List<SimpleScheduleData> allScheduleData, string excelFilePath, Action<int> progressCallback)
        {
            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            try
            {
                excel = new Excel.Application();
                workbook = excel.Workbooks.Add();

                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
                }

                Excel.Worksheet colorLegendSheet = (Excel.Worksheet)workbook.Worksheets[1];
                CreateColorLegendSheet(colorLegendSheet);

                int processedSchedules = 0;
                foreach (var scheduleData in allScheduleData)
                {
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                    worksheet.Name = scheduleData.Name.Length > 31 ? scheduleData.Name.Substring(0, 31) : scheduleData.Name;

                    ExportSimpleScheduleToWorksheet(scheduleData, worksheet);

                    processedSchedules++;
                    int percentage = (int)((double)processedSchedules / allScheduleData.Count * 100);
                    progressCallback?.Invoke(Math.Min(percentage, 100));
                }

                colorLegendSheet.Activate();

                excel.DisplayAlerts = false;
                workbook.SaveAs(excelFilePath);
            }
            finally
            {
                if (workbook != null) workbook.Close(false);
                if (excel != null) excel.Quit();
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        private void ExportSimpleScheduleToWorksheet(SimpleScheduleData scheduleData, Excel.Worksheet worksheet)
        {
            int currentColCount = Math.Max(1, scheduleData.Headers.Any() ? scheduleData.Headers.Count : (scheduleData.BodyRows.FirstOrDefault()?.Count ?? 1));
            int startRow = 1;

            // Title row
            Excel.Range titleRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, currentColCount]];
            titleRange.Merge();
            titleRange.Value2 = scheduleData.Name;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titleRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
            startRow++;

            // Column index row
            Excel.Range indexRow = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, currentColCount]];
            indexRow.Font.Bold = true;
            indexRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            indexRow.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
            indexRow.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            for (int col = 0; col < currentColCount; col++)
            {
                var cell = (Excel.Range)worksheet.Cells[startRow, col + 1];
                cell.Value2 = GetExcelColumnName(col + 1);
            }
            startRow++;

            // Headers with field information
            if (scheduleData.Headers.Any())
            {
                for (int col = 0; col < scheduleData.Headers.Count; col++)
                {
                    Excel.Range headerCell = (Excel.Range)worksheet.Cells[startRow, col + 1];
                    string headerText = scheduleData.Headers[col];

                    // Add field type information to header
                    if (col < scheduleData.FieldInfos.Count)
                    {
                        var fieldInfo = scheduleData.FieldInfos[col];
                        if (fieldInfo.IsCount)
                        {
                            headerText += "\n(Count)";
                        }
                        else if (fieldInfo.IsCalculatedField)
                        {
                            headerText += "\n(Calculated)";
                        }
                        else if (fieldInfo.IsSharedParameter)
                        {
                            headerText += "\n(Shared)";
                        }
                    }

                    headerCell.Value2 = headerText;
                    headerCell.Font.Bold = true;
                    headerCell.WrapText = true;
                    headerCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    // Apply special color for calculated/count fields
                    if (col < scheduleData.ColumnProperties.Count && scheduleData.ColumnProperties[col].IsCalculated)
                    {
                        headerCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#B19CD9")); // Light purple for calculated fields
                    }
                }
                ((Excel.Range)worksheet.Rows[startRow]).RowHeight = 35;
                startRow++;
            }

            // Separator row
            Excel.Range greyRowSeparator = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, currentColCount]];
            greyRowSeparator.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
            startRow++;

            // Body rows
            for (int row = 0; row < scheduleData.BodyRows.Count; row++)
            {
                var currentRowData = scheduleData.BodyRows[row];
                int currentRowInExcel = startRow + row;

                bool isBlankLine = row < scheduleData.IsBlankLineRow.Count && scheduleData.IsBlankLineRow[row];
                bool isHeaderOrFooter = row < scheduleData.IsGroupHeaderOrFooterRow.Count && scheduleData.IsGroupHeaderOrFooterRow[row];

                if (isBlankLine || isHeaderOrFooter)
                {
                    Excel.Range specialRowRange = worksheet.Range[worksheet.Cells[currentRowInExcel, 1], worksheet.Cells[currentRowInExcel, currentColCount]];
                    specialRowRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
                }

                for (int col = 0; col < currentRowData.Count; col++)
                {
                    Excel.Range dataCell = (Excel.Range)worksheet.Cells[currentRowInExcel, col + 1];
                    dataCell.Value2 = currentRowData[col];
                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    if (!isBlankLine && !isHeaderOrFooter)
                    {
                        if (col < scheduleData.ColumnProperties.Count)
                        {
                            var colProp = scheduleData.ColumnProperties[col];
                            if (colProp.IsCalculated)
                            {
                                // Light purple for calculated/count fields
                                dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E6E6FA"));
                            }
                            else if (colProp.IsReadOnly)
                            {
                                dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                            }
                            else if (colProp.IsType)
                            {
                                dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                            }
                        }
                    }
                }
            }

            // Add summary rows if present
            if (scheduleData.SummaryRows != null && scheduleData.SummaryRows.Any())
            {
                int summaryStartRow = startRow + scheduleData.BodyRows.Count;

                // Add separator before summary
                Excel.Range summarySeparator = worksheet.Range[worksheet.Cells[summaryStartRow, 1], worksheet.Cells[summaryStartRow, currentColCount]];
                summarySeparator.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
                summaryStartRow++;

                foreach (var summaryRow in scheduleData.SummaryRows)
                {
                    for (int col = 0; col < summaryRow.Count; col++)
                    {
                        Excel.Range summaryCell = (Excel.Range)worksheet.Cells[summaryStartRow, col + 1];
                        summaryCell.Value2 = summaryRow[col];
                        summaryCell.Font.Bold = true;
                        summaryCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        summaryCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFF8DC")); // Cream color for summary
                    }
                    summaryStartRow++;
                }
            }

            worksheet.Columns.AutoFit();
        }

        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int rem = columnNumber % 26;
                if (rem == 0)
                {
                    columnName = "Z" + columnName;
                    columnNumber = (columnNumber / 26) - 1;
                }
                else
                {
                    columnName = Convert.ToChar((rem - 1) + 'A') + columnName;
                    columnNumber = columnNumber / 26;
                }
            }
            return columnName;
        }

        // Import methods remain largely the same, but with awareness of calculated fields
        public List<ImportErrorItem> ImportSchedulesFromExcel(string excelFilePath, Action<int> progressCallback)
        {
            List<ImportErrorItem> errors = new List<ImportErrorItem>();
            Excel.Application excel = null;
            Excel.Workbook workbook = null;

            try
            {
                excel = new Excel.Application();
                workbook = excel.Workbooks.Open(excelFilePath);

                int totalSheets = workbook.Worksheets.Count;
                int processedSheets = 0;

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name == "Color Legend")
                    {
                        processedSheets++;
                        continue;
                    }

                    ViewSchedule schedule = FindScheduleByName(worksheet.Name);
                    if (schedule == null)
                    {
                        errors.Add(new ImportErrorItem { ElementId = "N/A", Description = $"Schedule '{worksheet.Name}' not found in model" });
                        processedSheets++;
                        continue;
                    }

                    var importErrors = ImportScheduleFromWorksheet(schedule, worksheet);
                    errors.AddRange(importErrors);

                    processedSheets++;
                    int percentage = (int)((double)processedSheets / totalSheets * 100);
                    progressCallback?.Invoke(percentage);
                }
            }
            finally
            {
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
            return errors;
        }

        private List<ImportErrorItem> ImportScheduleFromWorksheet(ViewSchedule schedule, Excel.Worksheet worksheet)
        {
            List<ImportErrorItem> errors = new List<ImportErrorItem>();
            ScheduleDefinition definition = schedule.Definition;

            // Get visible fields to understand which are editable
            var visibleFields = new List<ScheduleField>();
            var editableFieldIndices = new List<int>();

            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                if (!field.IsHidden)
                {
                    visibleFields.Add(field);
                    // Skip calculated fields and count fields during import
                    if (!field.IsCalculatedField && field.FieldType != ScheduleFieldType.Count)
                    {
                        editableFieldIndices.Add(visibleFields.Count - 1);
                    }
                }
            }

            Excel.Range usedRange = worksheet.UsedRange;
            int excelRows = usedRange.Rows.Count;
            int startRow = 5; // Skip title, index, header, and separator rows

            for (int i = startRow; i <= excelRows; i++)
            {
                try
                {
                    // Check if this is a special row (grey background)
                    var firstCell = usedRange.Cells[i, 1] as Excel.Range;
                    if (firstCell == null) continue;

                    // Skip rows with grey background (separators, headers, footers)
                    var cellColor = firstCell.Interior.Color;
                    if (cellColor != null)
                    {
                        int colorValue = Convert.ToInt32(cellColor);
                        if (colorValue == ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC")))
                        {
                            continue;
                        }
                    }

                    if (firstCell.Value2 == null) continue;
                    string keyValue = firstCell.Value2.ToString();

                    Element element = FindElementFromSchedule(schedule, keyValue, visibleFields);
                    if (element == null)
                    {
                        errors.Add(new ImportErrorItem { ElementId = keyValue, Description = "Element not found in schedule" });
                        continue;
                    }

                    for (int col = 0; col < editableFieldIndices.Count && col < usedRange.Columns.Count - 1; col++)
                    {
                        int fieldIndex = editableFieldIndices[col];
                        if (fieldIndex >= visibleFields.Count) continue;

                        var field = visibleFields[fieldIndex];

                        // Skip calculated and count fields
                        if (field.IsCalculatedField || field.FieldType == ScheduleFieldType.Count)
                            continue;

                        var dataCell = usedRange.Cells[i, col + 2] as Excel.Range; // +2 because first column is key
                        if (dataCell == null || dataCell.Value2 == null) continue;

                        string value = dataCell.Value2.ToString();
                        string paramName = field.GetName();

                        // Try to update the parameter
                        bool success = UpdateElementParameter(element, field, value);
                        if (!success)
                        {
                            // Try type parameter if instance failed
                            Element typeElement = _doc.GetElement(element.GetTypeId());
                            if (typeElement != null)
                            {
                                success = UpdateElementParameter(typeElement, field, value);
                            }
                        }

                        if (!success)
                        {
                            errors.Add(new ImportErrorItem
                            {
                                ElementId = element.Id.IntegerValue.ToString(),
                                Description = $"Failed to update parameter '{paramName}'"
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    errors.Add(new ImportErrorItem { ElementId = $"Row {i}", Description = ex.Message });
                }
            }
            return errors;
        }

        private bool UpdateElementParameter(Element element, ScheduleField field, string value)
        {
            if (element == null || field == null) return false;

            try
            {
                Parameter param = GetParameterByField(element, field);
                if (param != null && !param.IsReadOnly)
                {
                    return Utils.SetParameterValue(element, param.Definition.Name, value);
                }
            }
            catch { }

            return false;
        }

        private Element FindElementFromSchedule(ViewSchedule schedule, string keyValue, List<ScheduleField> visibleFields)
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc, schedule.Id);
            var elements = collector.ToElements();

            // Try to find by element ID first
            if (int.TryParse(keyValue, out int elementId))
            {
                Element element = _doc.GetElement(new ElementId(elementId));
                if (element != null && elements.Contains(element))
                    return element;
            }

            // Try to match by the first field value (usually a key field like Mark or Name)
            if (visibleFields.Any())
            {
                var firstField = visibleFields.First();
                string fieldName = firstField.GetName();

                foreach (var element in elements)
                {
                    Parameter param = GetParameterByField(element, firstField);
                    if (param != null)
                    {
                        string paramValue = Utils.GetParameterValue(element, fieldName);
                        if (paramValue == keyValue)
                            return element;
                    }
                }
            }

            return null;
        }

        private ViewSchedule FindScheduleByName(string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc);
            return collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(s => s.Name == name || (s.Name.Length > 31 && s.Name.Substring(0, 31) == name));
        }

        private void CreateColorLegendSheet(Excel.Worksheet colorLegendSheet)
        {
            colorLegendSheet.Name = "Color Legend";
            Excel.Range titleRange = colorLegendSheet.Range[colorLegendSheet.Cells[1, 2], colorLegendSheet.Cells[1, 4]];
            titleRange.Merge();
            titleRange.Value2 = "Color Legend";
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            titleRange.Borders.Weight = Excel.XlBorderWeight.xlThick;

            ((Excel.Range)colorLegendSheet.Cells[3, 2]).Value2 = "Color";
            ((Excel.Range)colorLegendSheet.Cells[3, 3]).Value2 = "Description";
            ((Excel.Range)colorLegendSheet.Cells[3, 4]).Value2 = "Notes";

            Excel.Range legendHeaderRange = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[3, 4]];
            legendHeaderRange.Font.Bold = true;
            legendHeaderRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            legendHeaderRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Existing color codes
            ((Excel.Range)colorLegendSheet.Cells[4, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
            ((Excel.Range)colorLegendSheet.Cells[4, 3]).Value2 = "Type value";
            ((Excel.Range)colorLegendSheet.Cells[4, 4]).Value2 = "Type parameters with the same ID should be filled the same";

            ((Excel.Range)colorLegendSheet.Cells[5, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
            ((Excel.Range)colorLegendSheet.Cells[5, 3]).Value2 = "Read-only value";
            ((Excel.Range)colorLegendSheet.Cells[5, 4]).Value2 = "Uneditable cell";

            ((Excel.Range)colorLegendSheet.Cells[6, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
            ((Excel.Range)colorLegendSheet.Cells[6, 3]).Value2 = "Parameter does not exist for element";
            ((Excel.Range)colorLegendSheet.Cells[6, 4]).Value2 = "Applies to Category export only";

            ((Excel.Range)colorLegendSheet.Cells[7, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
            ((Excel.Range)colorLegendSheet.Cells[7, 3]).Value2 = "Title / Main Header Row";
            ((Excel.Range)colorLegendSheet.Cells[7, 4]).Value2 = "Indicates a title or header row";

            ((Excel.Range)colorLegendSheet.Cells[8, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
            ((Excel.Range)colorLegendSheet.Cells[8, 3]).Value2 = "Separator / Index / Group Header or Blank Line";
            ((Excel.Range)colorLegendSheet.Cells[8, 4]).Value2 = "Indicates a separator, index row, or a schedule group header/footer/blank line";

            // New color codes for schedule-specific fields
            ((Excel.Range)colorLegendSheet.Cells[9, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#B19CD9"));
            ((Excel.Range)colorLegendSheet.Cells[9, 3]).Value2 = "Calculated Field Header";
            ((Excel.Range)colorLegendSheet.Cells[9, 4]).Value2 = "Header for calculated or count fields in schedules";

            ((Excel.Range)colorLegendSheet.Cells[10, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E6E6FA"));
            ((Excel.Range)colorLegendSheet.Cells[10, 3]).Value2 = "Calculated Field Value";
            ((Excel.Range)colorLegendSheet.Cells[10, 4]).Value2 = "Values from calculated or count fields (read-only)";

            ((Excel.Range)colorLegendSheet.Cells[11, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFF8DC"));
            ((Excel.Range)colorLegendSheet.Cells[11, 3]).Value2 = "Summary/Grand Total Row";
            ((Excel.Range)colorLegendSheet.Cells[11, 4]).Value2 = "Summary or grand total rows from schedules";

            Excel.Range dataRange = colorLegendSheet.Range[colorLegendSheet.Cells[4, 2], colorLegendSheet.Cells[11, 4]];
            dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range entireTable = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[11, 4]];
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            ((Excel.Range)colorLegendSheet.Columns[2]).ColumnWidth = 15;
            ((Excel.Range)colorLegendSheet.Columns[3]).ColumnWidth = 35;
            ((Excel.Range)colorLegendSheet.Columns[4]).ColumnWidth = 50;
        }
    }

    public class ScheduleItem : INotifyPropertyChanged
    {
        private ViewSchedule _schedule;
        private bool _isSelected;
        private string _scheduleName;
        private bool _isSelectAll;

        public ViewSchedule Schedule { get => _schedule; set { _schedule = value; OnPropertyChanged(nameof(Schedule)); } }
        public string ScheduleName { get => _scheduleName; set { _scheduleName = value; OnPropertyChanged(nameof(ScheduleName)); } }
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }
        public bool IsSelectAll { get => _isSelectAll; set { _isSelectAll = value; OnPropertyChanged(nameof(IsSelectAll)); OnPropertyChanged(nameof(FontWeight)); OnPropertyChanged(nameof(TextColor)); } }
        public string FontWeight => IsSelectAll ? "Bold" : "Normal";
        public string TextColor => IsSelectAll ? "#000000" : "#000000";

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public ScheduleItem(ViewSchedule schedule)
        {
            Schedule = schedule;
            ScheduleName = schedule.Name;
        }
        public ScheduleItem(string displayName, bool isSelectAll = false)
        {
            ScheduleName = displayName;
            IsSelectAll = isSelectAll;
        }
    }
}