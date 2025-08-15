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
        public bool IsCalculated { get; set; }
        public string FieldType { get; set; }
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
        public List<ScheduleFieldInfo> FieldInfos { get; set; }
        public List<List<bool>> ParameterExistsForRow { get; set; }
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
            var collector = new FilteredElementCollector(_doc);
            var schedules = collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTemplate && !s.IsTitleblockRevisionSchedule && s.Definition != null)
                .OrderBy(s => s.Name)
                .ToList();
            return schedules;
        }

        public List<SimpleScheduleData> GetScheduleDataForExport(List<ViewSchedule> schedules, bool includeHeaders)
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
                    FieldInfos = new List<ScheduleFieldInfo>(),
                    ParameterExistsForRow = new List<List<bool>>()
                };

                ScheduleDefinition definition = schedule.Definition;
                TableData tableData = schedule.GetTableData();
                TableSectionData bodySection = tableData.GetSectionData(SectionType.Body);

                var visibleFields = new List<ScheduleField>();
                for (int i = 0; i < definition.GetFieldCount(); i++)
                {
                    var field = definition.GetField(i);
                    if (!field.IsHidden) visibleFields.Add(field);
                }

                int numberOfRows = bodySection.NumberOfRows;
                var elementsInSchedule = new FilteredElementCollector(_doc, schedule.Id).WhereElementIsNotElementType().ToList();

                if (includeHeaders)
                {
                    for (int colIdx = 0; colIdx < visibleFields.Count; colIdx++)
                    {
                        var field = visibleFields[colIdx];
                        simpleData.Headers.Add(field.GetName());
                        simpleData.ColumnLetters.Add(GetExcelColumnName(colIdx + 1));
                        var fieldInfo = GetScheduleFieldInfo(field, definition);
                        simpleData.FieldInfos.Add(fieldInfo);

                        var colProp = new ColumnProperty
                        {
                            IsCalculated = fieldInfo.IsCalculatedField || fieldInfo.IsCount
                        };

                        if (!colProp.IsCalculated)
                        {
                            bool isEverWritable = false, isEverInstance = false, isEverType = false;
                            foreach (var element in elementsInSchedule)
                            {
                                if (GetParameterByField(element, field) is Parameter instParam)
                                {
                                    isEverInstance = true;
                                    if (!instParam.IsReadOnly) isEverWritable = true;
                                }
                                else if (_doc.GetElement(element.GetTypeId()) is Element typeElement && GetParameterByField(typeElement, field) is Parameter typeParam)
                                {
                                    isEverType = true;
                                    if (!typeParam.IsReadOnly) isEverWritable = true;
                                }
                            }
                            colProp.IsReadOnly = !isEverWritable;
                            colProp.IsType = isEverType && !isEverInstance;
                            if (fieldInfo.IsSharedParameter) colProp.IsReadOnly = false;
                        }
                        simpleData.ColumnProperties.Add(colProp);
                    }
                }

                int bodyStartRow = (definition.ShowHeaders && numberOfRows > 0) ? 1 : 0;
                for (int row = bodyStartRow; row < numberOfRows; row++)
                {
                    var rowData = new List<string>();
                    for (int col = 0; col < simpleData.Headers.Count; col++)
                    {
                        rowData.Add(schedule.GetCellText(SectionType.Body, row, col));
                    }
                    simpleData.BodyRows.Add(rowData);
                }

                foreach (var rowData in simpleData.BodyRows)
                {
                    bool isBlank = rowData.All(string.IsNullOrWhiteSpace);
                    simpleData.IsBlankLineRow.Add(isBlank);
                    if (isBlank)
                    {
                        simpleData.IsGroupHeaderOrFooterRow.Add(false);
                        continue;
                    }
                    var firstCellText = rowData.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c));
                    int nonEmptyCount = rowData.Count(c => !string.IsNullOrWhiteSpace(c));
                    bool isHeaderOrFooter = (firstCellText != null && (firstCellText.ToLower().Contains("total") || firstCellText.Contains(":"))) || nonEmptyCount == 1;
                    simpleData.IsGroupHeaderOrFooterRow.Add(isHeaderOrFooter);
                }

                int elementIndex = 0;
                for (int i = 0; i < simpleData.BodyRows.Count; i++)
                {
                    var parameterExistsList = new List<bool>(new bool[simpleData.Headers.Count]);
                    if (!simpleData.IsBlankLineRow[i] && !simpleData.IsGroupHeaderOrFooterRow[i] && elementIndex < elementsInSchedule.Count)
                    {
                        Element elementForRow = elementsInSchedule[elementIndex];
                        for (int col = 0; col < simpleData.Headers.Count; col++)
                        {
                            var field = visibleFields[col];
                            var fieldInfo = simpleData.FieldInfos[col];

                            if (fieldInfo.IsCalculatedField || fieldInfo.IsCount)
                            {
                                parameterExistsList[col] = true;
                            }
                            else
                            {
                                Parameter param = GetParameterByField(elementForRow, field);
                                if (param == null && _doc.GetElement(elementForRow.GetTypeId()) is Element typeElem)
                                {
                                    param = GetParameterByField(typeElem, field);
                                }
                                if (param != null)
                                {
                                    parameterExistsList[col] = true;
                                }
                            }
                        }
                        elementIndex++;
                    }
                    simpleData.ParameterExistsForRow.Add(parameterExistsList);
                }

                if (definition.ShowGrandTotal)
                {
                    var summarySection = tableData.GetSectionData(SectionType.Summary);
                    if (summarySection != null)
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

        public ScheduleFieldInfo GetScheduleFieldInfo(ScheduleField field, ScheduleDefinition definition)
        {
            var fieldInfo = new ScheduleFieldInfo
            {
                Name = field.GetName(),
                FieldType = field.FieldType,
                ParameterId = field.ParameterId,
                IsCount = field.FieldType == ScheduleFieldType.Count,
                IsCalculatedField = field.IsCalculatedField
            };

            if (!fieldInfo.IsCalculatedField && field.ParameterId != null && field.ParameterId != ElementId.InvalidElementId)
            {
                try
                {
                    if (_doc.GetElement(field.ParameterId) is SharedParameterElement)
                    {
                        fieldInfo.IsSharedParameter = true;
                    }
                }
                catch { }
            }
            return fieldInfo;
        }

        public Parameter GetParameterByField(Element element, ScheduleField field)
        {
            if (element == null || field == null) return null;

            if (field.ParameterId != null && field.ParameterId.IntegerValue != ElementId.InvalidElementId.IntegerValue)
            {
                // Check if it's a BuiltInParameter
                if (Enum.IsDefined(typeof(BuiltInParameter), field.ParameterId.IntegerValue))
                {
                    var bip = (BuiltInParameter)field.ParameterId.IntegerValue;
                    if (element.get_Parameter(bip) is Parameter p) return p;
                }
                // Check if it's a shared/project parameter by iterating
                else
                {
                    foreach (Parameter p in element.Parameters)
                    {
                        if (p.Id.Equals(field.ParameterId)) return p;
                    }
                }
            }

            // Fallback for cases where ParameterId might not be reliable
            if (element.LookupParameter(field.GetName()) is Parameter pByName) return pByName;

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

            Excel.Range titleRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, currentColCount]];
            titleRange.Merge();
            titleRange.Value2 = scheduleData.Name;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titleRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
            startRow++;

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

            if (scheduleData.Headers.Any())
            {
                Excel.Range headerRowRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, currentColCount]];
                headerRowRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
                headerRowRange.Font.Bold = true;
                headerRowRange.WrapText = false;
                headerRowRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                headerRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                for (int col = 0; col < scheduleData.Headers.Count; col++)
                {
                    Excel.Range headerCell = (Excel.Range)worksheet.Cells[startRow, col + 1];
                    headerCell.Value2 = scheduleData.Headers[col];
                }
                ((Excel.Range)worksheet.Rows[startRow]).RowHeight = 25;
                startRow++;
            }

            for (int row = 0; row < scheduleData.BodyRows.Count; row++)
            {
                var currentRowData = scheduleData.BodyRows[row];
                int currentRowInExcel = startRow + row;

                bool isBlankLine = scheduleData.IsBlankLineRow[row];
                bool isHeaderOrFooter = scheduleData.IsGroupHeaderOrFooterRow[row];

                for (int col = 0; col < currentRowData.Count; col++)
                {
                    Excel.Range dataCell = (Excel.Range)worksheet.Cells[currentRowInExcel, col + 1];
                    dataCell.Value2 = currentRowData[col];
                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    if (isBlankLine || isHeaderOrFooter)
                    {
                        dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
                        if (currentRowData.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c))?.ToLower().Contains("total") == true)
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFF8DC"));
                            dataCell.Font.Bold = true;
                        }
                    }
                    else
                    {
                        bool parameterExists = scheduleData.ParameterExistsForRow.Count > row && scheduleData.ParameterExistsForRow[row].Count > col && scheduleData.ParameterExistsForRow[row][col];

                        if (!parameterExists)
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                        }
                        else if (col < scheduleData.ColumnProperties.Count)
                        {
                            var colProp = scheduleData.ColumnProperties[col];
                            if (colProp.IsCalculated)
                            {
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

            if (scheduleData.SummaryRows.Any())
            {
                int summaryStartRow = startRow + scheduleData.BodyRows.Count;
                worksheet.Range[worksheet.Cells[summaryStartRow, 1], worksheet.Cells[summaryStartRow, currentColCount]].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
                summaryStartRow++;

                foreach (var summaryRow in scheduleData.SummaryRows)
                {
                    Excel.Range summaryRowRange = worksheet.Range[worksheet.Cells[summaryStartRow, 1], worksheet.Cells[summaryStartRow, currentColCount]];
                    summaryRowRange.Font.Bold = true;
                    summaryRowRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    summaryRowRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFF8DC"));
                    for (int col = 0; col < summaryRow.Count; col++)
                    {
                        ((Excel.Range)worksheet.Cells[summaryStartRow, col + 1]).Value2 = summaryRow[col];
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

        public List<ImportErrorItem> ImportSchedulesFromExcel(string excelFilePath, Action<int> progressCallback)
        {
            var errors = new List<ImportErrorItem>();
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
                    if (FindScheduleByName(worksheet.Name) is ViewSchedule schedule)
                    {
                        errors.AddRange(ImportScheduleFromWorksheet(schedule, worksheet));
                    }
                    else
                    {
                        errors.Add(new ImportErrorItem { ElementId = "N/A", Description = $"Schedule '{worksheet.Name}' not found in model" });
                    }
                    processedSheets++;
                    progressCallback?.Invoke((int)((double)processedSheets / totalSheets * 100));
                }
            }
            finally
            {
                if (workbook != null) workbook.Close(false);
                if (excel != null) excel.Quit();
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
            return errors;
        }

        private List<ImportErrorItem> ImportScheduleFromWorksheet(ViewSchedule schedule, Excel.Worksheet worksheet)
        {
            var errors = new List<ImportErrorItem>();
            var definition = schedule.Definition;
            var visibleFields = new List<ScheduleField>();
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                if (!definition.GetField(i).IsHidden) visibleFields.Add(definition.GetField(i));
            }

            Excel.Range usedRange = worksheet.UsedRange;
            for (int i = 4; i <= usedRange.Rows.Count; i++)
            {
                try
                {
                    var firstCell = usedRange.Cells[i, 1] as Excel.Range;
                    int cellColor = Convert.ToInt32(firstCell?.Interior.Color);
                    if (cellColor == ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC")) ||
                        cellColor == ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFF8DC")))
                    {
                        continue;
                    }

                    string keyValue = firstCell?.Value2?.ToString();
                    if (string.IsNullOrEmpty(keyValue)) continue;

                    if (FindElementFromSchedule(schedule, keyValue, visibleFields) is Element element)
                    {
                        for (int col = 0; col < visibleFields.Count && col < usedRange.Columns.Count; col++)
                        {
                            var field = visibleFields[col];
                            if (field.IsCalculatedField || field.FieldType == ScheduleFieldType.Count) continue;

                            var dataCell = usedRange.Cells[i, col + 1] as Excel.Range;
                            string value = dataCell?.Value2?.ToString();
                            if (value == null) continue;

                            if (!UpdateElementParameter(element, field, value))
                            {
                                errors.Add(new ImportErrorItem
                                {
                                    ElementId = element.Id.IntegerValue.ToString(),
                                    Description = $"Failed to update parameter '{field.GetName()}'"
                                });
                            }
                        }
                    }
                    else
                    {
                        errors.Add(new ImportErrorItem { ElementId = keyValue, Description = "Element not found in schedule" });
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
                if (GetParameterByField(element, field) is Parameter param && !param.IsReadOnly)
                {
                    return Utils.SetParameterValue(element, param.Definition.Name, value);
                }
                if (_doc.GetElement(element.GetTypeId()) is Element typeElem && GetParameterByField(typeElem, field) is Parameter typeParam && !typeParam.IsReadOnly)
                {
                    return Utils.SetParameterValue(typeElem, typeParam.Definition.Name, value);
                }
            }
            catch { }
            return false;
        }

        private Element FindElementFromSchedule(ViewSchedule schedule, string keyValue, List<ScheduleField> visibleFields)
        {
            var elements = new FilteredElementCollector(_doc, schedule.Id).ToElements();
            if (int.TryParse(keyValue, out int elementId))
            {
                if (elements.FirstOrDefault(e => e.Id.IntegerValue == elementId) is Element elem) return elem;
            }
            return null;
        }

        private ViewSchedule FindScheduleByName(string name)
        {
            return new FilteredElementCollector(_doc)
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

            int row = 4;
            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Type value";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Type parameters with the same ID should be filled the same";

            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Read-only value";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Uneditable cell";

            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Parameter does not exist for element";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Applies to Category export only";

            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Title / Main Header Row";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Indicates a title or header row";

            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CCCCCC"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Separator / Index / Group Header or Blank Line";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Indicates a separator, index row, or a schedule group header/footer/blank line";

            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E6E6FA"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Calculated Field Value";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Values from calculated or count fields (read-only)";

            ((Excel.Range)colorLegendSheet.Cells[row, 2]).Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFF8DC"));
            ((Excel.Range)colorLegendSheet.Cells[row, 3]).Value2 = "Summary/Grand Total Row";
            ((Excel.Range)colorLegendSheet.Cells[row++, 4]).Value2 = "Summary or grand total rows from schedules";

            Excel.Range dataRange = colorLegendSheet.Range[colorLegendSheet.Cells[4, 2], colorLegendSheet.Cells[row - 1, 4]];
            dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range entireTable = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[row - 1, 4]];
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            colorLegendSheet.Columns.AutoFit();
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