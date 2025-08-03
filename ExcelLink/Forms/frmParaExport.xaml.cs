using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using ExcelLink.Common;
using System.Reflection;

namespace ExcelLink.Forms
{
    /// <summary>
    /// Interaction with frmParaExport.xaml
    /// </summary>
    public partial class frmParaExport : Window, INotifyPropertyChanged
    {
        private Document _doc;
        private ObservableCollection<ParaExportCategoryItem> _categoryItems;
        private ObservableCollection<ParaExportParameterItem> _availableParameterItems;
        private ObservableCollection<ParaExportParameterItem> _selectedParameterItems;

        public ObservableCollection<ParaExportCategoryItem> CategoryItems
        {
            get { return _categoryItems; }
            set
            {
                _categoryItems = value;
                OnPropertyChanged(nameof(CategoryItems));
            }
        }

        public ObservableCollection<ParaExportParameterItem> AvailableParameterItems
        {
            get { return _availableParameterItems; }
            set
            {
                _availableParameterItems = value;
                OnPropertyChanged(nameof(AvailableParameterItems));
            }
        }

        public ObservableCollection<ParaExportParameterItem> SelectedParameterItems
        {
            get { return _selectedParameterItems; }
            set
            {
                _selectedParameterItems = value;
                OnPropertyChanged(nameof(SelectedParameterItems));
            }
        }

        public bool IsEntireModelChecked
        {
            get { return (bool)rbEntireModel.IsChecked; }
        }

        public bool IsActiveViewChecked
        {
            get { return (bool)rbActiveView.IsChecked; }
        }

        public List<string> SelectedCategoryNames
        {
            get
            {
                return CategoryItems
                    .Where(item => item.IsSelected && !item.IsSelectAll)
                    .Select(item => item.CategoryName)
                    .ToList();
            }
        }

        public List<string> SelectedParameterNames
        {
            get
            {
                return SelectedParameterItems
                    .Select(item => item.ParameterName)
                    .ToList();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public frmParaExport(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            DataContext = this;

            // Initialize collections
            CategoryItems = new ObservableCollection<ParaExportCategoryItem>();
            AvailableParameterItems = new ObservableCollection<ParaExportParameterItem>();
            SelectedParameterItems = new ObservableCollection<ParaExportParameterItem>();

            // Load initial data
            LoadCategoriesBasedOnScope();
            lvAvailableParameters.ItemsSource = AvailableParameterItems;
            lvSelectedParameters.ItemsSource = SelectedParameterItems;
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            ImportFromExcel();
        }

        private void ExportToExcel()
        {
            // Validate selections
            if (!SelectedCategoryNames.Any())
            {
                TaskDialog.Show("Error", "Please select at least one category.");
                return;
            }

            if (!SelectedParameterNames.Any())
            {
                TaskDialog.Show("Error", "Please select at least one parameter.");
                return;
            }

            // Prompt user to save Excel file
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel files|*.xlsx";
            saveDialog.Title = "Save Revit Parameters to Excel";

            string defaultFileName = _doc.Title;
            if (string.IsNullOrEmpty(defaultFileName))
            {
                defaultFileName = "RevitParameterExport";
            }
            saveDialog.FileName = defaultFileName + ".xlsx";

            if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string excelFile = saveDialog.FileName;

            // Show progress bar
            ShowProgressBar();

            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Worksheet colorLegendSheet = null;

            try
            {
                // Create Excel application and workbook
                excel = new Excel.Application();
                workbook = excel.Workbooks.Add();

                // Remove default sheets except the first one
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
                }

                // Create the Color Legend Sheet first
                colorLegendSheet = (Excel.Worksheet)workbook.Worksheets[1];
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

                // Write legend content
                Excel.Range greyCell = (Excel.Range)colorLegendSheet.Cells[4, 2];
                greyCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                ((Excel.Range)colorLegendSheet.Cells[4, 3]).Value2 = "Parameter does not exist for this element";
                ((Excel.Range)colorLegendSheet.Cells[4, 4]).Value2 = "Do not fill or edit cell";

                Excel.Range lightYellowCell = (Excel.Range)colorLegendSheet.Cells[5, 2];
                lightYellowCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                ((Excel.Range)colorLegendSheet.Cells[5, 3]).Value2 = "Type value";
                ((Excel.Range)colorLegendSheet.Cells[5, 4]).Value2 = "Type parameters with the same ID should be filled the same";

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

                // Get elements based on selected categories and scope
                var selectedCategories = CategoryItems
                    .Where(item => item.IsSelected && !item.IsSelectAll)
                    .ToList();

                // Process each category
                int sheetIndex = 1;
                foreach (var categoryItem in selectedCategories)
                {
                    sheetIndex++;
                    // Create or get worksheet
                    if (workbook.Worksheets.Count < sheetIndex)
                    {
                        worksheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                    }
                    else
                    {
                        worksheet = (Excel.Worksheet)workbook.Worksheets[sheetIndex];
                    }

                    // Set sheet name (Excel limits sheet names to 31 characters)
                    string sheetName = categoryItem.CategoryName.Length > 31 ?
                        categoryItem.CategoryName.Substring(0, 31) : categoryItem.CategoryName;
                    worksheet.Name = sheetName;

                    // Write headers with multi-line text
                    Excel.Range elementIdHeader = (Excel.Range)worksheet.Cells[1, 1];
                    elementIdHeader.Value2 = "Element ID";
                    elementIdHeader.ColumnWidth = 12;

                    List<string> selectedParameters = SelectedParameterNames;

                    for (int i = 0; i < selectedParameters.Count; i++)
                    {
                        string paramName = selectedParameters[i];
                        string paramType = "N/A";
                        string paramStorageType = "N/A";

                        // Get an instance element to check parameters
                        FilteredElementCollector tempCollector;
                        if (IsEntireModelChecked)
                        {
                            tempCollector = new FilteredElementCollector(_doc);
                        }
                        else
                        {
                            tempCollector = new FilteredElementCollector(_doc, _doc.ActiveView.Id);
                        }

                        tempCollector.OfCategoryId(categoryItem.Category.Id);
                        tempCollector.WhereElementIsNotElementType();
                        Element tempElement = tempCollector.FirstElement();

                        Parameter param = null;
                        bool isTypeParam = false;

                        if (tempElement != null)
                        {
                            // First check instance parameter
                            param = tempElement.LookupParameter(paramName);
                            if (param != null)
                            {
                                paramType = "Instance Parameter";
                            }
                            else
                            {
                                // Check type parameter
                                Element typeElem = _doc.GetElement(tempElement.GetTypeId());
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
                                BuiltInParameter bip = Utils.GetBuiltInParameterByName(paramName);
                                if (bip != BuiltInParameter.INVALID)
                                {
                                    param = tempElement.get_Parameter(bip);
                                    if (param != null)
                                    {
                                        paramType = "Instance Parameter";
                                    }
                                    else
                                    {
                                        // Check if it's a built-in type parameter
                                        Element typeElem = _doc.GetElement(tempElement.GetTypeId());
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
                                paramStorageType = Utils.GetParameterStorageTypeString(param.StorageType);
                            }
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

                    // Add auto-filter to the headers
                    headerRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

                    // Set row height to accommodate 3 lines of text
                    ((Excel.Range)worksheet.Rows[1]).RowHeight = 45;

                    // Get all elements in the category and scope for data writing
                    FilteredElementCollector dataCollector;
                    if (IsEntireModelChecked)
                    {
                        dataCollector = new FilteredElementCollector(_doc);
                    }
                    else
                    {
                        dataCollector = new FilteredElementCollector(_doc, _doc.ActiveView.Id);
                    }
                    dataCollector.OfCategoryId(categoryItem.Category.Id);
                    dataCollector.WhereElementIsNotElementType();
                    List<Element> elements = dataCollector.ToList();

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
                                value = Utils.GetParameterValue(element, paramName);
                            }
                            else
                            {
                                // Check if the parameter exists as a type parameter
                                Element typeElem = _doc.GetElement(element.GetTypeId());
                                if (typeElem != null)
                                {
                                    param = typeElem.LookupParameter(paramName);
                                    if (param != null)
                                    {
                                        value = Utils.GetParameterValue(typeElem, paramName);
                                        isTypeParam = true;
                                    }
                                }
                            }

                            // If still not found, check built-in parameters
                            if (param == null)
                            {
                                BuiltInParameter bip = Utils.GetBuiltInParameterByName(paramName);
                                if (bip != BuiltInParameter.INVALID)
                                {
                                    param = element.get_Parameter(bip);
                                    if (param != null)
                                    {
                                        value = Utils.GetParameterValue(element, paramName);
                                    }
                                    else
                                    {
                                        // Check if it's a built-in type parameter
                                        Element typeElem = _doc.GetElement(element.GetTypeId());
                                        if (typeElem != null)
                                        {
                                            param = typeElem.get_Parameter(bip);
                                            if (param != null)
                                            {
                                                value = Utils.GetParameterValue(typeElem, paramName);
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

                    // Auto-fit columns
                    worksheet.Columns.AutoFit();

                    UpdateProgressBar((int)((double)sheetIndex / (selectedCategories.Count + 1) * 100));
                }

                // Save the file
                workbook.SaveAs(excelFile);

                // Activate the Color Legend sheet and make Excel visible
                colorLegendSheet.Activate();
                excel.Visible = true;

                TaskDialog.Show("Success", "Export completed successfully!");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to export parameters:\n{ex.Message}");
            }
            finally
            {
                HideProgressBar();
                // We do not close the workbook and quit the application in the finally block
                // as the user wants the file to remain open. We only release the COM objects.
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (colorLegendSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(colorLegendSheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        private void ImportFromExcel()
        {
            // Prompt user to select Excel file
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel files|*.xlsx;*.xls";
            openDialog.Title = "Select Excel File to Import";

            if (openDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string excelFile = openDialog.FileName;

            // Show progress bar
            ShowProgressBar();

            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;

            try
            {
                // Create Excel application
                excel = new Excel.Application();
                workbook = excel.Workbooks.Open(excelFile);

                // Find the correct worksheet to import from, ignoring the "Color Legend" sheet
                worksheet = workbook.Worksheets.Cast<Excel.Worksheet>()
                                    .FirstOrDefault(s => s.Name != "Color Legend");

                if (worksheet == null)
                {
                    TaskDialog.Show("Error", "Could not find a valid worksheet to import from.");
                    return;
                }

                usedRange = worksheet.UsedRange;

                // Get headers
                List<string> headers = new List<string>();
                for (int j = 1; j <= usedRange.Columns.Count; j++)
                {
                    var headerCell = usedRange.Cells[1, j] as Excel.Range;
                    if (headerCell != null && headerCell.Value2 != null)
                    {
                        headers.Add(headerCell.Value2.ToString());
                    }
                }

                int elementIdIndex = headers.IndexOf("ElementId");
                if (elementIdIndex == -1)
                {
                    TaskDialog.Show("Error", "The Excel file must contain a column named 'ElementId'.");
                    return;
                }

                // Track errors for a summary
                List<string> errorMessages = new List<string>();

                // Start Revit transaction
                using (Transaction t = new Transaction(_doc, "Import Parameters from Excel"))
                {
                    t.Start();
                    // Loop through rows
                    for (int i = 2; i <= usedRange.Rows.Count; i++)
                    {
                        var idCell = usedRange.Cells[i, elementIdIndex + 1] as Excel.Range;

                        // Handle potential non-numeric ElementId
                        if (idCell == null || idCell.Value2 == null) continue;

                        string idString = idCell.Value2.ToString();
                        int elementIdInt;

                        if (!int.TryParse(idString, out elementIdInt))
                        {
                            errorMessages.Add($"Row {i}: Failed to parse ElementId '{idString}'. Skipping row.");
                            continue;
                        }

                        ElementId elementId = new ElementId(elementIdInt);
                        Element element = _doc.GetElement(elementId);

                        if (element != null)
                        {
                            // Loop through parameters
                            for (int j = 0; j < headers.Count; j++)
                            {
                                if (j != elementIdIndex)
                                {
                                    string paramName = headers[j];
                                    var paramCell = usedRange.Cells[i, j + 1] as Excel.Range;
                                    string paramValue = paramCell?.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(paramValue))
                                    {
                                        try
                                        {
                                            Utils.SetParameterValue(element, paramName, paramValue);
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Row {i}: Failed to set parameter '{paramName}' with value '{paramValue}'. Error: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            errorMessages.Add($"Row {i}: Element with ID '{elementIdInt}' not found in model. Skipping row.");
                        }
                        UpdateProgressBar((int)((double)(i - 1) / (usedRange.Rows.Count - 1) * 100));
                    }
                    t.Commit();
                }

                if (errorMessages.Any())
                {
                    string summary = "Import completed with errors:\n" + string.Join("\n", errorMessages.Take(10));
                    if (errorMessages.Count > 10)
                    {
                        summary += $"\n...and {errorMessages.Count - 10} more errors.";
                    }
                    TaskDialog.Show("Import Completed with Errors", summary);
                }
                else
                {
                    TaskDialog.Show("Success", "Import completed successfully!");
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to import parameters:\n{ex.Message}");
            }
            finally
            {
                HideProgressBar();
                if (workbook != null) workbook.Close(false);
                if (excel != null) excel.Quit();
                if (usedRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        public void UpdateProgressBar(int percentage)
        {
            progressBar.Value = percentage;
            progressBarText.Text = $"{percentage}%";
        }

        public void ShowProgressBar()
        {
            progressBar.Visibility = System.Windows.Visibility.Visible;
            progressBarText.Visibility = System.Windows.Visibility.Visible;
        }

        public void HideProgressBar()
        {
            progressBar.Visibility = System.Windows.Visibility.Collapsed;
            progressBarText.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void rbEntireModel_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded)
            {
                LoadCategoriesBasedOnScope();
            }
        }

        private void rbActiveView_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded)
            {
                LoadCategoriesBasedOnScope();
            }
        }

        private void LoadCategoriesBasedOnScope()
        {
            CategoryItems.Clear();
            AvailableParameterItems.Clear();
            SelectedParameterItems.Clear();

            List<Category> availableCategories = GetCategoriesWithElementsInScope();

            // Add "Select All" option
            CategoryItems.Add(new ParaExportCategoryItem("Select All Categories", true));

            // Add individual categories
            foreach (Category category in availableCategories.OrderBy(c => c.Name))
            {
                CategoryItems.Add(new ParaExportCategoryItem(category));
            }

            // Set ListView source
            lvCategories.ItemsSource = CategoryItems;

            // Initialize search box
            txtCategorySearch.Text = "Search categories...";
        }

        private List<Category> GetCategoriesWithElementsInScope()
        {
            FilteredElementCollector collector;

            if (IsEntireModelChecked)
            {
                collector = new FilteredElementCollector(_doc);
            }
            else
            {
                collector = new FilteredElementCollector(_doc, _doc.ActiveView.Id);
            }

            // Get all element instances
            var elementInstances = collector
                .WhereElementIsNotElementType()
                .ToList();

            // Get unique categories that have element instances
            var categoriesWithElements = elementInstances
                .Where(e => e.Category != null)
                .Select(e => e.Category)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .ToList();

            // List of category names to exclude
            HashSet<string> excludedCategoryNames = new HashSet<string>
            {
                "Survey Point",
                "Sun Path",
                "Project Information",
                "Project Base Point",
                "Primary Contours",
                "Material Assets",
                "Legend Components",
                "Internal Origin",
                "Cameras",
                "HVAC Zones",
                "Pipe Segments",
                "Area Based Load Type",
                "Circuit Naming Scheme",
                "<Sketch>",
                "Center Line",
                "Center line", // Different casing
                "Lines",
                "Detail Items",
                "Model Lines",
                "Detail Lines",
                "<Room Separation>",
                "<Area Boundary>",
                "<Space Separation>",
                "Curtain Panel Tags", // Exclude tags
                "Curtain System Tags",
                "Detail Item Tags",
                "Door Tags",
                "Floor Tags",
                "Generic Annotations",
                "Keynote Tags",
                "Material Tags",
                "Multi-Category Tags",
                "Parking Tags",
                "Plumbing Fixture Tags",
                "Property Line Segment Tags",
                "Property Tags",
                "Revision Clouds",
                "Room Tags",
                "Space Tags",
                "Structural Annotations",
                "Wall Tags",
                "Window Tags"
            };

            // Filter to only include model categories AND Rooms (which is not a model category)
            var modelCategories = categoriesWithElements
                .Where(c => (c.CategoryType == CategoryType.Model || c.Name == "Rooms") &&
                           !excludedCategoryNames.Contains(c.Name) &&
                           !c.Name.ToLower().Contains("line") && // Exclude any category with "line" in the name
                           !c.Name.ToLower().Contains("sketch")) // Exclude any category with "sketch" in the name
                .ToList();

            return modelCategories;
        }

        private void CategoryCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox checkBox && checkBox.DataContext is ParaExportCategoryItem categoryItem)
            {
                if (categoryItem.IsSelectAll)
                {
                    // Handle "Select All" checkbox
                    bool isChecked = checkBox.IsChecked == true;
                    foreach (ParaExportCategoryItem item in CategoryItems)
                    {
                        if (!item.IsSelectAll)
                        {
                            item.IsSelected = isChecked;
                        }
                    }
                }
                else
                {
                    // Handle individual category checkbox
                    UpdateCategorySelectAllCheckboxState();
                }

                UpdateCategorySearchTextBox();
                LoadParametersForSelectedCategories();
            }
        }

        private void UpdateCategorySelectAllCheckboxState()
        {
            var selectAllItem = CategoryItems.FirstOrDefault(item => item.IsSelectAll);
            if (selectAllItem != null)
            {
                var categoryItems = CategoryItems.Where(item => !item.IsSelectAll).ToList();
                int selectedCount = categoryItems.Count(item => item.IsSelected);
                int totalCount = categoryItems.Count;

                if (selectedCount == 0)
                {
                    selectAllItem.IsSelected = false;
                }
                else if (selectedCount == totalCount)
                {
                    selectAllItem.IsSelected = true;
                }
            }
        }

        private void UpdateCategorySearchTextBox()
        {
            var selectedItems = CategoryItems.Where(item => item.IsSelected && !item.IsSelectAll).ToList();

            if (selectedItems.Count == 0)
            {
                txtCategorySearch.Text = "Search categories...";
            }
            else if (selectedItems.Count == 1)
            {
                txtCategorySearch.Text = selectedItems.First().CategoryName;
            }
            else
            {
                txtCategorySearch.Text = $"{selectedItems.Count} categories selected";
            }
        }

        private void LoadParametersForSelectedCategories()
        {
            AvailableParameterItems.Clear();
            SelectedParameterItems.Clear();

            var selectedCategories = CategoryItems
                .Where(item => item.IsSelected && !item.IsSelectAll)
                .ToList();

            if (!selectedCategories.Any())
            {
                return;
            }

            HashSet<string> allParameterNames = new HashSet<string>();
            Dictionary<string, Parameter> parameterMap = new Dictionary<string, Parameter>();

            foreach (var categoryItem in selectedCategories)
            {
                if (categoryItem.Category != null)
                {
                    FilteredElementCollector collector;

                    if (IsEntireModelChecked)
                    {
                        collector = new FilteredElementCollector(_doc);
                    }
                    else
                    {
                        collector = new FilteredElementCollector(_doc, _doc.ActiveView.Id);
                    }

                    // Get element instances in the category
                    collector.OfCategoryId(categoryItem.Category.Id);
                    collector.WhereElementIsNotElementType();

                    var instances = collector.ToList();

                    if (!instances.Any()) continue;

                    // Collect parameters from instances
                    foreach (Element instance in instances)
                    {
                        // Get all parameters from the element (including built-in)
                        foreach (Parameter param in instance.Parameters)
                        {
                            if (param != null && param.Definition != null)
                            {
                                string paramName = param.Definition.Name;
                                if (!allParameterNames.Contains(paramName))
                                {
                                    allParameterNames.Add(paramName);
                                    parameterMap[paramName] = param;
                                }
                            }
                        }

                        // Get type parameters
                        ElementId typeId = instance.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element elementType = _doc.GetElement(typeId);
                            if (elementType != null)
                            {
                                foreach (Parameter param in elementType.Parameters)
                                {
                                    if (param != null && param.Definition != null)
                                    {
                                        string paramName = param.Definition.Name;
                                        if (!allParameterNames.Contains(paramName))
                                        {
                                            allParameterNames.Add(paramName);
                                            parameterMap[paramName] = param;
                                        }
                                    }
                                }
                            }
                        }

                        // Add important built-in parameters that might not show up in Parameters collection
                        AddSpecificBuiltInParameters(instance, allParameterNames, parameterMap);
                    }
                }
            }

            // List of parameter names to exclude
            HashSet<string> excludedParameters = new HashSet<string>
            {
                "Phase Created",
                "Phase Demolished",
                "View Template",
                "Design Option",
                "Edited by",
                "View Scale",
                "Detail Level",
                "Visible",
                "Graphics Overrides",
                "Family Name"
            };

            // List of ElementId parameters that are allowed (exception list)
            HashSet<string> allowedElementIdParameters = new HashSet<string>
            {
                "Family",
                "Family and Type",
                "Workset"
            };

            // Filter and sort parameters
            var distinctParameters = parameterMap
                .Where(kvp => !excludedParameters.Contains(kvp.Key) &&
                             ((kvp.Value.StorageType == StorageType.String ||
                               kvp.Value.StorageType == StorageType.Double ||
                               kvp.Value.StorageType == StorageType.Integer) ||
                              (kvp.Value.StorageType == StorageType.ElementId &&
                               allowedElementIdParameters.Contains(kvp.Key))) &&
                             kvp.Value.StorageType != StorageType.None)
                .OrderBy(kvp => kvp.Key)
                .ToList();

            // Populate the list of all available parameters
            foreach (var kvp in distinctParameters)
            {
                bool isTypeParam = false;
                if (kvp.Value.Element is ElementType)
                {
                    isTypeParam = true;
                }

                AvailableParameterItems.Add(new ParaExportParameterItem(kvp.Value, kvp.Value.IsReadOnly, isTypeParam));
            }
        }

        private void AddSpecificBuiltInParameters(Element element, HashSet<string> parameterNames, Dictionary<string, Parameter> parameterMap)
        {
            var builtInParamsToCheck = new Dictionary<string, BuiltInParameter>
            {
                { "Family", BuiltInParameter.ELEM_FAMILY_PARAM },
                { "Family and Type", BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM },
                { "Type", BuiltInParameter.ELEM_TYPE_PARAM },
                { "Type Name", BuiltInParameter.SYMBOL_NAME_PARAM },
                { "Comments", BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS },
                { "Type Comments", BuiltInParameter.ALL_MODEL_TYPE_COMMENTS },
                { "Mark", BuiltInParameter.ALL_MODEL_MARK },
                { "Type Mark", BuiltInParameter.ALL_MODEL_TYPE_MARK },
                { "Description", BuiltInParameter.ALL_MODEL_DESCRIPTION },
                { "Manufacturer", BuiltInParameter.ALL_MODEL_MANUFACTURER },
                { "Model", BuiltInParameter.ALL_MODEL_MODEL },
                { "URL", BuiltInParameter.ALL_MODEL_URL },
                { "Cost", BuiltInParameter.ALL_MODEL_COST },
                { "Assembly Code", BuiltInParameter.UNIFORMAT_CODE },
                { "Assembly Description", BuiltInParameter.UNIFORMAT_DESCRIPTION },
                { "Keynote", BuiltInParameter.KEYNOTE_PARAM },
            };

            if (element.Category != null)
            {
                string categoryName = element.Category.Name;

                if (categoryName.Contains("Floor"))
                {
                    builtInParamsToCheck["Default Thickness"] = BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM;
                    builtInParamsToCheck["Thickness"] = BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM;
                    builtInParamsToCheck["Function"] = BuiltInParameter.FUNCTION_PARAM;
                    builtInParamsToCheck["Structural"] = BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;
                }
                else if (categoryName.Contains("Wall"))
                {
                    builtInParamsToCheck["Width"] = BuiltInParameter.WALL_ATTR_WIDTH_PARAM;
                    builtInParamsToCheck["Function"] = BuiltInParameter.FUNCTION_PARAM;
                    builtInParamsToCheck["Height"] = BuiltInParameter.WALL_USER_HEIGHT_PARAM;
                    builtInParamsToCheck["Base Offset"] = BuiltInParameter.WALL_BASE_OFFSET;
                    builtInParamsToCheck["Top Offset"] = BuiltInParameter.WALL_TOP_OFFSET;
                }
                else if (categoryName.Contains("Door") || categoryName.Contains("Window"))
                {
                    builtInParamsToCheck["Head Height"] = BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM;
                    builtInParamsToCheck["Sill Height"] = BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM;
                }

                builtInParamsToCheck["Area"] = BuiltInParameter.HOST_AREA_COMPUTED;
                builtInParamsToCheck["Volume"] = BuiltInParameter.HOST_VOLUME_COMPUTED;
                builtInParamsToCheck["Perimeter"] = BuiltInParameter.HOST_PERIMETER_COMPUTED;
                builtInParamsToCheck["Level"] = BuiltInParameter.LEVEL_PARAM;
            }

            foreach (var kvp in builtInParamsToCheck)
            {
                try
                {
                    Parameter param = element.get_Parameter(kvp.Value);
                    if (param == null)
                    {
                        ElementId typeId = element.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element elementType = _doc.GetElement(typeId);
                            if (elementType != null)
                            {
                                param = elementType.get_Parameter(kvp.Value);
                            }
                        }
                    }

                    if (param != null && param.Definition != null && !parameterNames.Contains(kvp.Key))
                    {
                        parameterNames.Add(kvp.Key);
                        parameterMap[kvp.Key] = param;
                    }
                }
                catch { }
            }
        }

        private void ParameterCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            // The logic for moving parameters is now handled by the new move buttons.
        }

        // Search functionality for categories
        private void txtCategorySearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.IsFocused)
            {
                string searchText = textBox.Text.ToLower();
                if (searchText == "search categories...") return;

                var filteredItems = CategoryItems.Where(c => c.IsSelectAll || c.CategoryName.ToLower().Contains(searchText));
                lvCategories.ItemsSource = new ObservableCollection<ParaExportCategoryItem>(filteredItems);
            }
        }

        private void txtCategorySearch_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.Text == "Search categories...")
            {
                textBox.Text = "";
            }
        }

        private void txtCategorySearch_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                txtCategorySearch.Text = "Search categories...";
                lvCategories.ItemsSource = CategoryItems;
            }
        }

        // Search functionality for parameters
        private void txtParameterSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.IsFocused)
            {
                string searchText = textBox.Text.ToLower();

                if (searchText == "Search parameters...")
                {
                    lvAvailableParameters.ItemsSource = AvailableParameterItems;
                    return;
                }

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    lvAvailableParameters.ItemsSource = AvailableParameterItems;
                }
                else
                {
                    var filteredParameters = AvailableParameterItems.Where(p => p.ParameterName.ToLower().Contains(searchText));
                    lvAvailableParameters.ItemsSource = new ObservableCollection<ParaExportParameterItem>(filteredParameters);
                }
            }
        }

        private void txtParameterSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.Text == "Search parameters...")
            {
                textBox.Text = "";
            }
        }

        private void txtParameterSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                lvAvailableParameters.ItemsSource = AvailableParameterItems;
                textBox.Text = "Search parameters...";
            }
        }

        // New methods for double-click functionality
        private void lvAvailableParameters_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lvAvailableParameters.SelectedItem is ParaExportParameterItem item)
            {
                MoveParameter(item, AvailableParameterItems, SelectedParameterItems);
            }
        }

        private void lvSelectedParameters_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lvSelectedParameters.SelectedItem is ParaExportParameterItem item)
            {
                MoveParameter(item, SelectedParameterItems, AvailableParameterItems);
            }
        }

        private void MoveParameter(ParaExportParameterItem item, ObservableCollection<ParaExportParameterItem> source, ObservableCollection<ParaExportParameterItem> destination)
        {
            if (item != null)
            {
                source.Remove(item);
                destination.Add(item);
            }
        }

        // New button event handlers for moving items
        private void btnMoveRight_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvAvailableParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();

            foreach (var item in selectedItems)
            {
                if (AvailableParameterItems.Contains(item))
                {
                    AvailableParameterItems.Remove(item);
                    SelectedParameterItems.Add(item);
                }
            }
        }

        private void btnMoveLeft_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();

            foreach (var item in selectedItems)
            {
                if (SelectedParameterItems.Contains(item))
                {
                    SelectedParameterItems.Remove(item);
                    AvailableParameterItems.Add(item);
                }
            }
        }

        private void btnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.Items.Cast<ParaExportParameterItem>().ToList();
            if (selectedItems.Count == 0) return;

            foreach (var item in selectedItems)
            {
                int index = SelectedParameterItems.IndexOf(item);
                if (index > 0)
                {
                    SelectedParameterItems.Move(index, index - 1);
                }
            }
        }

        private void btnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.Items.Cast<ParaExportParameterItem>().ToList();
            if (selectedItems.Count == 0) return;

            for (int i = selectedItems.Count - 1; i >= 0; i--)
            {
                var item = selectedItems[i];
                int index = SelectedParameterItems.IndexOf(item);
                if (index < SelectedParameterItems.Count - 1)
                {
                    SelectedParameterItems.Move(index, index + 1);
                }
            }
        }

        private void btnMoveAllRight_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in AvailableParameterItems.ToList())
            {
                AvailableParameterItems.Remove(item);
                SelectedParameterItems.Add(item);
            }
        }

        private void btnMoveAllLeft_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in SelectedParameterItems.ToList())
            {
                SelectedParameterItems.Remove(item);
                AvailableParameterItems.Add(item);
            }
        }
    }

    // Helper classes for data binding
    public class ParaExportCategoryItem : INotifyPropertyChanged
    {
        private Category _category;
        private bool _isSelected;
        private string _categoryName;
        private bool _isSelectAll;

        public Category Category
        {
            get { return _category; }
            set
            {
                _category = value;
                OnPropertyChanged(nameof(Category));
            }
        }

        public string CategoryName
        {
            get { return _categoryName; }
            set
            {
                _categoryName = value;
                OnPropertyChanged(nameof(CategoryName));
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

        // Constructor for regular categories
        public ParaExportCategoryItem(Category category)
        {
            Category = category;
            CategoryName = category.Name;
            IsSelected = false;
            IsSelectAll = false;
        }

        // Constructor for "Select All" item
        public ParaExportCategoryItem(string displayName, bool isSelectAll = false)
        {
            Category = null;
            CategoryName = displayName;
            IsSelected = false;
            IsSelectAll = isSelectAll;
        }
    }

    public class ParaExportParameterItem : INotifyPropertyChanged
    {
        private Parameter _parameter;
        private bool _isSelected;
        private string _parameterName;
        private bool _isSelectAll;
        private SolidColorBrush _parameterColor;

        public Parameter Parameter
        {
            get { return _parameter; }
            set
            {
                _parameter = value;
                OnPropertyChanged(nameof(Parameter));
            }
        }

        public string ParameterName
        {
            get { return _parameterName; }
            set
            {
                _parameterName = value;
                OnPropertyChanged(nameof(ParameterName));
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

        public SolidColorBrush ParameterColor
        {
            get { return _parameterColor; }
            set
            {
                _parameterColor = value;
                OnPropertyChanged(nameof(ParameterColor));
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

        // Constructor for regular parameters
        public ParaExportParameterItem(Parameter parameter, bool isReadOnly, bool isTypeParam)
        {
            Parameter = parameter;
            ParameterName = parameter.Definition.Name;
            IsSelected = false;
            IsSelectAll = false;

            // Set color based on parameter properties
            if (isReadOnly)
            {
                // Read-only parameter color
                ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FF4747"));
            }
            else if (isTypeParam)
            {
                // Type parameter color
                ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FFE699"));
            }
            else
            {
                // Editable instance parameter (default)
                ParameterColor = new SolidColorBrush(Colors.White);
            }
        }

        // Constructor for "Select All" item
        public ParaExportParameterItem(string displayName, bool isSelectAll = false)
        {
            Parameter = null;
            ParameterName = displayName;
            IsSelected = false;
            IsSelectAll = isSelectAll;
            ParameterColor = new SolidColorBrush(Colors.White); // Default color for "Select All"
        }
    }
}