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
    }

    /// <summary>
    /// A simple data class to hold schedule information, free of any Revit API objects.
    /// This allows it to be safely passed to a background thread.
    /// </summary>
    public class SimpleScheduleData
    {
        public string Name { get; set; }
        public List<string> Headers { get; set; }
        public List<string> ColumnLetters { get; set; }
        public List<ColumnProperty> ColumnProperties { get; set; }
        public List<List<string>> BodyRows { get; set; }
        public List<List<string>> SummaryRows { get; set; }
    }

    public class ScheduleManager
    {
        private Document _doc;

        public ScheduleManager(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// Gets all schedules in the document
        /// </summary>
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

        /// <summary>
        /// Reads data from Revit schedules and stores it in a simple data structure.
        /// This method MUST be called from the main Revit thread.
        /// </summary>
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
                    SummaryRows = new List<List<string>>()
                };

                ScheduleDefinition definition = schedule.Definition;
                TableData tableData = schedule.GetTableData();
                TableSectionData bodySection = tableData.GetSectionData(SectionType.Body);
                int numberOfRows = bodySection.NumberOfRows;
                int numberOfColumns = bodySection.NumberOfColumns;

                // Get a sample element to check parameter properties
                Element sampleElement = new FilteredElementCollector(_doc, schedule.Id).FirstElement();
                Element sampleTypeElement = sampleElement != null ? _doc.GetElement(sampleElement.GetTypeId()) : null;

                // Get Column Headers, Letters, and Properties
                if (includeHeaders)
                {
                    TableSectionData headerSection = tableData.GetSectionData(SectionType.Header);
                    if (headerSection != null && headerSection.NumberOfRows > 0)
                    {
                        for (int col = 0; col < numberOfColumns; col++)
                        {
                            simpleData.Headers.Add(schedule.GetCellText(SectionType.Header, 0, col));
                            simpleData.ColumnLetters.Add(GetExcelColumnName(col + 1));

                            var field = definition.GetField(col);
                            string paramName = field.GetName();
                            var colProp = new ColumnProperty { IsReadOnly = true, IsType = false }; // Default to read-only for calculated/special fields

                            if (sampleElement != null)
                            {
                                Parameter param = sampleElement.LookupParameter(paramName);
                                if (param != null)
                                {
                                    colProp.IsReadOnly = param.IsReadOnly;
                                    colProp.IsType = false;
                                }
                                else if (sampleTypeElement != null)
                                {
                                    param = sampleTypeElement.LookupParameter(paramName);
                                    if (param != null)
                                    {
                                        colProp.IsReadOnly = param.IsReadOnly;
                                        colProp.IsType = true;
                                    }
                                }
                            }
                            simpleData.ColumnProperties.Add(colProp);
                        }
                    }
                }

                // Read body rows
                for (int row = 0; row < numberOfRows; row++)
                {
                    var rowData = new List<string>();
                    for (int col = 0; col < numberOfColumns; col++)
                    {
                        rowData.Add(schedule.GetCellText(SectionType.Body, row, col));
                    }
                    simpleData.BodyRows.Add(rowData);
                }

                // Read grand totals (summary)
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
        /// Exports the pre-processed simple schedule data to Excel.
        /// This method contains no Revit API calls and is safe to run on a background thread.
        /// </summary>
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

                excel.DisplayAlerts = false;
                workbook.SaveAs(excelFilePath);
            }
            finally
            {
                if (workbook != null) workbook.Close(false);
                if (excel != null) excel.Quit();
                // Release COM objects
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        /// <summary>
        /// Helper method to write data from a SimpleScheduleData object to a worksheet.
        /// </summary>
        private void ExportSimpleScheduleToWorksheet(SimpleScheduleData scheduleData, Excel.Worksheet worksheet)
        {
            int currentColCount = Math.Max(1, scheduleData.Headers.Any() ? scheduleData.Headers.Count : (scheduleData.BodyRows.FirstOrDefault()?.Count ?? 1));
            int startRow = 1;

            // Merged Title Row
            Excel.Range titleRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, currentColCount]];
            titleRange.Merge();
            titleRange.Value2 = scheduleData.Name;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titleRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E0E0E0"));
            startRow++;

            // Header Row
            if (scheduleData.Headers.Any())
            {
                for (int col = 0; col < scheduleData.Headers.Count; col++)
                {
                    Excel.Range headerCell = (Excel.Range)worksheet.Cells[startRow, col + 1];
                    headerCell.Value2 = scheduleData.Headers[col];
                    headerCell.Font.Bold = true;
                    headerCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
                    headerCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
                startRow++;
            }

            // Body Data Rows
            for (int row = 0; row < scheduleData.BodyRows.Count; row++)
            {
                var currentRowData = scheduleData.BodyRows[row];
                bool isEmptyRow = currentRowData.All(string.IsNullOrWhiteSpace);
                int currentRowInExcel = startRow + row;

                for (int col = 0; col < currentRowData.Count; col++)
                {
                    Excel.Range dataCell = (Excel.Range)worksheet.Cells[currentRowInExcel, col + 1];
                    dataCell.Value2 = currentRowData[col];
                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    if (isEmptyRow)
                    {
                        Excel.Range rowRange = worksheet.Range[worksheet.Cells[currentRowInExcel, 1], worksheet.Cells[currentRowInExcel, currentColCount]];
                        rowRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D0D0D0"));
                    }
                    else if (col < scheduleData.ColumnProperties.Count)
                    {
                        var colProp = scheduleData.ColumnProperties[col];
                        if (colProp.IsReadOnly) dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                        else if (colProp.IsType) dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                    }
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


        /// <summary>
        /// Imports schedules from Excel
        /// </summary>
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

        /// <summary>
        /// Imports data from worksheet to schedule
        /// </summary>
        private List<ImportErrorItem> ImportScheduleFromWorksheet(ViewSchedule schedule, Excel.Worksheet worksheet)
        {
            List<ImportErrorItem> errors = new List<ImportErrorItem>();
            ScheduleDefinition definition = schedule.Definition;
            IList<ScheduleFieldId> fieldIds = definition.GetFieldOrder();
            Excel.Range usedRange = worksheet.UsedRange;
            int excelRows = usedRange.Rows.Count;
            int startRow = 3; // Start from row 3 to skip Title and Headers

            for (int i = startRow; i <= excelRows; i++)
            {
                try
                {
                    var firstCell = usedRange.Cells[i, 1] as Excel.Range;
                    if (firstCell == null || firstCell.Value2 == null) continue;
                    string keyValue = firstCell.Value2.ToString();

                    Element element = FindElementFromSchedule(schedule, keyValue);
                    if (element == null)
                    {
                        errors.Add(new ImportErrorItem { ElementId = keyValue, Description = "Element not found in model" });
                        continue;
                    }

                    for (int col = 2; col <= usedRange.Columns.Count; col++)
                    {
                        var dataCell = usedRange.Cells[i, col] as Excel.Range;
                        if (dataCell == null || dataCell.Value2 == null) continue;
                        string value = dataCell.Value2.ToString();

                        if (col - 1 < fieldIds.Count)
                        {
                            ScheduleField field = definition.GetField(fieldIds[col - 1]);
                            string paramName = field.GetName();

                            bool success = Utils.SetParameterValue(element, paramName, value);
                            if (!success)
                            {
                                errors.Add(new ImportErrorItem { ElementId = element.Id.IntegerValue.ToString(), Description = $"Failed to update parameter '{paramName}'" });
                            }
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

        /// <summary>
        /// Finds schedule by name
        /// </summary>
        private ViewSchedule FindScheduleByName(string name)
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc);
            return collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(s => s.Name == name || (s.Name.Length > 31 && s.Name.Substring(0, 31) == name));
        }

        /// <summary>
        /// Finds element from schedule based on key value
        /// </summary>
        private Element FindElementFromSchedule(ViewSchedule schedule, string keyValue)
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc, schedule.Id);
            var elements = collector.ToElements();

            if (int.TryParse(keyValue, out int elementId))
            {
                Element element = _doc.GetElement(new ElementId(elementId));
                if (element != null) return element;
            }

            foreach (var element in elements)
            {
                if (element.Name == keyValue) return element;
                Parameter markParam = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                if (markParam != null && markParam.AsString() == keyValue) return element;
            }
            return null;
        }

        /// <summary>
        /// Creates the color legend sheet
        /// </summary>
        private void CreateColorLegendSheet(Excel.Worksheet colorLegendSheet)
        {
            colorLegendSheet.Name = "Color Legend";
            Excel.Range titleRange = colorLegendSheet.Range[colorLegendSheet.Cells[1, 2], colorLegendSheet.Cells[1, 4]];
            titleRange.Merge();
            titleRange.Value2 = "Schedule Export Color Legend";
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

            Excel.Range yellowCell = (Excel.Range)colorLegendSheet.Cells[4, 2];
            yellowCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
            ((Excel.Range)colorLegendSheet.Cells[4, 3]).Value2 = "Header Row";
            ((Excel.Range)colorLegendSheet.Cells[4, 4]).Value2 = "Column headers from schedule";

            Excel.Range greyCell = (Excel.Range)colorLegendSheet.Cells[5, 2];
            greyCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E0E0E0"));
            ((Excel.Range)colorLegendSheet.Cells[5, 3]).Value2 = "Summary/Total Row";
            ((Excel.Range)colorLegendSheet.Cells[5, 4]).Value2 = "Grand totals or summary data";

            Excel.Range whiteCell = (Excel.Range)colorLegendSheet.Cells[6, 2];
            whiteCell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
            ((Excel.Range)colorLegendSheet.Cells[6, 3]).Value2 = "Data Row";
            ((Excel.Range)colorLegendSheet.Cells[6, 4]).Value2 = "Regular schedule data";

            Excel.Range dataRange = colorLegendSheet.Range[colorLegendSheet.Cells[4, 2], colorLegendSheet.Cells[6, 4]];
            dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range entireTable = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[6, 4]];
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            ((Excel.Range)colorLegendSheet.Columns[2]).ColumnWidth = 15;
            ((Excel.Range)colorLegendSheet.Columns[3]).ColumnWidth = 30;
            ((Excel.Range)colorLegendSheet.Columns[4]).ColumnWidth = 40;
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