using Autodesk.Revit.DB;
using ExcelLink.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLink.Common
{
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
        /// Exports selected schedules to Excel
        /// </summary>
        public void ExportSchedulesToExcel(List<ViewSchedule> schedules, string excelFilePath,
            Action<int> progressCallback, bool includeHeaders = true, bool includeGrandTotals = true)
        {
            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Worksheet colorLegendSheet = null;

            try
            {
                excel = new Excel.Application();
                workbook = excel.Workbooks.Add();

                // Remove default sheets except the first one
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
                }

                // Create Color Legend sheet
                colorLegendSheet = (Excel.Worksheet)workbook.Worksheets[1];
                CreateColorLegendSheet(colorLegendSheet);

                int totalSchedules = schedules.Count;
                int processedSchedules = 0;

                foreach (var schedule in schedules)
                {
                    // Create new worksheet for each schedule
                    worksheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);

                    // Limit sheet name to 31 characters
                    string sheetName = schedule.Name.Length > 31 ?
                        schedule.Name.Substring(0, 31) : schedule.Name;
                    worksheet.Name = sheetName;

                    // Export schedule data
                    ExportScheduleToWorksheet(schedule, worksheet, includeHeaders, includeGrandTotals);

                    processedSchedules++;
                    int percentage = (int)((double)processedSchedules / totalSchedules * 100);
                    progressCallback?.Invoke(Math.Min(percentage, 100));
                }

                // Save workbook
                excel.DisplayAlerts = false;
                workbook.SaveAs(excelFilePath);
                excel.DisplayAlerts = true;

                // Activate Color Legend sheet
                colorLegendSheet.Activate();
            }
            finally
            {
                // Clean up COM objects
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (colorLegendSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(colorLegendSheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        /// <summary>
        /// Exports a single schedule to a worksheet
        /// </summary>
        private void ExportScheduleToWorksheet(ViewSchedule schedule, Excel.Worksheet worksheet,
            bool includeHeaders, bool includeGrandTotals)
        {
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

            int numberOfRows = sectionData.NumberOfRows;
            int numberOfColumns = sectionData.NumberOfColumns;

            // Write headers if requested
            int startRow = 1;
            if (includeHeaders)
            {
                TableSectionData headerData = tableData.GetSectionData(SectionType.Header);
                if (headerData != null && headerData.NumberOfRows > 0)
                {
                    for (int col = 0; col < numberOfColumns; col++)
                    {
                        string headerText = schedule.GetCellText(SectionType.Header, 0, col);
                        Excel.Range headerCell = (Excel.Range)worksheet.Cells[startRow, col + 1];
                        headerCell.Value2 = headerText;
                        headerCell.Font.Bold = true;
                        headerCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
                        headerCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    startRow++;
                }
            }

            // Write body data
            for (int row = 0; row < numberOfRows; row++)
            {
                for (int col = 0; col < numberOfColumns; col++)
                {
                    string cellText = schedule.GetCellText(SectionType.Body, row, col);
                    Excel.Range dataCell = (Excel.Range)worksheet.Cells[startRow + row, col + 1];
                    dataCell.Value2 = cellText;
                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                }
            }

            // Write grand totals if requested and available
            if (includeGrandTotals)
            {
                TableSectionData summaryData = tableData.GetSectionData(SectionType.Summary);
                if (summaryData != null && summaryData.NumberOfRows > 0)
                {
                    int summaryStartRow = startRow + numberOfRows;
                    for (int row = 0; row < summaryData.NumberOfRows; row++)
                    {
                        for (int col = 0; col < summaryData.NumberOfColumns; col++)
                        {
                            string cellText = schedule.GetCellText(SectionType.Summary, row, col);
                            Excel.Range summaryCell = (Excel.Range)worksheet.Cells[summaryStartRow + row, col + 1];
                            summaryCell.Value2 = cellText;
                            summaryCell.Font.Bold = true;
                            summaryCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E0E0E0"));
                            summaryCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                    }
                }
            }

            // Auto-fit columns
            worksheet.Columns.AutoFit();
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
                    // Skip Color Legend sheet
                    if (worksheet.Name == "Color Legend")
                    {
                        processedSheets++;
                        continue;
                    }

                    // Find matching schedule in Revit
                    ViewSchedule schedule = FindScheduleByName(worksheet.Name);
                    if (schedule == null)
                    {
                        errors.Add(new ImportErrorItem
                        {
                            ElementId = "N/A",
                            Description = $"Schedule '{worksheet.Name}' not found in model"
                        });
                        processedSheets++;
                        continue;
                    }

                    // Import data from worksheet to schedule
                    var importErrors = ImportScheduleFromWorksheet(schedule, worksheet);
                    errors.AddRange(importErrors);

                    processedSheets++;
                    int percentage = (int)((double)processedSheets / totalSheets * 100);
                    progressCallback?.Invoke(percentage);
                }
            }
            finally
            {
                // Clean up COM objects
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

            // Note: Revit schedules are generally read-only from the API perspective
            // We can only modify the underlying elements that appear in the schedule

            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

            // Get schedule fields to understand what parameters we're dealing with
            ScheduleDefinition definition = schedule.Definition;
            IList<ScheduleFieldId> fieldIds = definition.GetFieldOrder();

            Excel.Range usedRange = worksheet.UsedRange;
            int excelRows = usedRange.Rows.Count;
            int startRow = 2; // Assuming row 1 is header

            // For each row in Excel, try to find and update the corresponding element
            for (int i = startRow; i <= excelRows; i++)
            {
                try
                {
                    // Get the first column value (usually element ID or key field)
                    var firstCell = usedRange.Cells[i, 1] as Excel.Range;
                    if (firstCell == null || firstCell.Value2 == null) continue;

                    string keyValue = firstCell.Value2.ToString();

                    // Find element in model based on schedule type and key value
                    Element element = FindElementFromSchedule(schedule, keyValue);
                    if (element == null)
                    {
                        errors.Add(new ImportErrorItem
                        {
                            ElementId = keyValue,
                            Description = "Element not found in model"
                        });
                        continue;
                    }

                    // Update element parameters based on Excel data
                    for (int col = 2; col <= usedRange.Columns.Count; col++)
                    {
                        var dataCell = usedRange.Cells[i, col] as Excel.Range;
                        if (dataCell == null || dataCell.Value2 == null) continue;

                        string value = dataCell.Value2.ToString();

                        // Get corresponding field from schedule
                        if (col - 1 < fieldIds.Count)
                        {
                            ScheduleField field = definition.GetField(fieldIds[col - 1]);
                            string paramName = field.GetName();

                            // Try to set parameter value
                            bool success = Utils.SetParameterValue(element, paramName, value);
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
                }
                catch (Exception ex)
                {
                    errors.Add(new ImportErrorItem
                    {
                        ElementId = $"Row {i}",
                        Description = ex.Message
                    });
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
            // Get elements in schedule
            FilteredElementCollector collector = new FilteredElementCollector(_doc, schedule.Id);
            var elements = collector.ToElements();

            // Try to find by ID first
            if (int.TryParse(keyValue, out int elementId))
            {
                Element element = _doc.GetElement(new ElementId(elementId));
                if (element != null) return element;
            }

            // Try to find by name or mark
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

            // Merge and center title
            Excel.Range titleRange = colorLegendSheet.Range[colorLegendSheet.Cells[1, 2], colorLegendSheet.Cells[1, 4]];
            titleRange.Merge();
            titleRange.Value2 = "Schedule Export Color Legend";
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;

            // Add borders to title
            titleRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            titleRange.Borders.Weight = Excel.XlBorderWeight.xlThick;

            // Write legend headers
            ((Excel.Range)colorLegendSheet.Cells[3, 2]).Value2 = "Color";
            ((Excel.Range)colorLegendSheet.Cells[3, 3]).Value2 = "Description";
            ((Excel.Range)colorLegendSheet.Cells[3, 4]).Value2 = "Notes";

            // Format headers
            Excel.Range legendHeaderRange = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[3, 4]];
            legendHeaderRange.Font.Bold = true;
            legendHeaderRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            legendHeaderRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Write legend content
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

            // Apply borders
            Excel.Range dataRange = colorLegendSheet.Range[colorLegendSheet.Cells[4, 2], colorLegendSheet.Cells[6, 4]];
            dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            // Apply thick outside border
            Excel.Range entireTable = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[6, 4]];
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            // Set column widths
            ((Excel.Range)colorLegendSheet.Columns[2]).ColumnWidth = 15;
            ((Excel.Range)colorLegendSheet.Columns[3]).ColumnWidth = 30;
            ((Excel.Range)colorLegendSheet.Columns[4]).ColumnWidth = 40;
        }
    }

    /// <summary>
    /// Schedule item for UI binding
    /// </summary>
    public class ScheduleItem : INotifyPropertyChanged
    {
        private ViewSchedule _schedule;
        private bool _isSelected;
        private string _scheduleName;
        private bool _isSelectAll;

        public ViewSchedule Schedule
        {
            get { return _schedule; }
            set
            {
                _schedule = value;
                OnPropertyChanged(nameof(Schedule));
            }
        }

        public string ScheduleName
        {
            get { return _scheduleName; }
            set
            {
                _scheduleName = value;
                OnPropertyChanged(nameof(ScheduleName));
            }
        }

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public bool IsSelectAll
        {
            get { return _isSelectAll; }
            set
            {
                _isSelectAll = value;
                OnPropertyChanged(nameof(IsSelectAll));
                OnPropertyChanged(nameof(FontWeight));
                OnPropertyChanged(nameof(TextColor));
            }
        }

        public string FontWeight
        {
            get { return IsSelectAll ? "Bold" : "Normal"; }
        }

        public string TextColor
        {
            get { return IsSelectAll ? "#000000" : "#000000"; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ScheduleItem(ViewSchedule schedule)
        {
            Schedule = schedule;
            ScheduleName = schedule.Name;
            IsSelected = false;
            IsSelectAll = false;
        }

        public ScheduleItem(string displayName, bool isSelectAll = false)
        {
            Schedule = null;
            ScheduleName = displayName;
            IsSelected = false;
            IsSelectAll = isSelectAll;
        }
    }
}