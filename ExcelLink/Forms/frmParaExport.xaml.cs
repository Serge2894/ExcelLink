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
using System.Globalization;
using System.Windows.Threading;

namespace ExcelLink.Forms
{
    /// <summary>
    /// Interaction with frmParaExport.xaml
    /// </summary>
    public partial class frmParaExport : Window, INotifyPropertyChanged
    {
        private Document _doc;
        private ExternalEvent _importExternalEvent;
        private ImportEventHandler _importEventHandler;
        private ObservableCollection<ParaExportCategoryItem> _categoryItems;
        private ObservableCollection<ParaExportParameterItem> _availableParameterItems;
        private ObservableCollection<ParaExportParameterItem> _selectedParameterItems;
        private ObservableCollection<ScheduleItem> _scheduleItems;
        private ScheduleManager _scheduleManager;
        private bool _isMaximized = false;
        private DispatcherTimer _progressTimer;
        private int _targetProgress;
        private int _currentProgress;
        private System.Windows.Controls.TabControl _mainTabControl;

        #region Properties

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

        public ObservableCollection<ScheduleItem> ScheduleItems
        {
            get { return _scheduleItems; }
            set
            {
                _scheduleItems = value;
                OnPropertyChanged(nameof(ScheduleItems));
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

        public List<ViewSchedule> SelectedSchedules
        {
            get
            {
                return ScheduleItems
                    .Where(item => item.IsSelected && !item.IsSelectAll && item.Schedule != null)
                    .Select(item => item.Schedule)
                    .ToList();
            }
        }

        #endregion

        #region Constructor and Initialization

        public frmParaExport(Document doc, ExternalEvent importExternalEvent, ImportEventHandler importEventHandler)
        {
            InitializeComponent();
            _doc = doc;
            _importExternalEvent = importExternalEvent;
            _importEventHandler = importEventHandler;
            _scheduleManager = new ScheduleManager(doc);
            DataContext = this;

            // Initialize collections
            CategoryItems = new ObservableCollection<ParaExportCategoryItem>();
            AvailableParameterItems = new ObservableCollection<ParaExportParameterItem>();
            SelectedParameterItems = new ObservableCollection<ParaExportParameterItem>();
            ScheduleItems = new ObservableCollection<ScheduleItem>();

            // Initialize progress timer for smooth animation
            _progressTimer = new DispatcherTimer();
            _progressTimer.Interval = TimeSpan.FromMilliseconds(20);
            _progressTimer.Tick += ProgressTimer_Tick;

            // Find the TabControl
            _mainTabControl = this.FindName("mainTabControl") as System.Windows.Controls.TabControl;

            // Load initial data
            LoadCategoriesBasedOnScope();
            LoadSchedules();

            lvAvailableParameters.ItemsSource = AvailableParameterItems;
            lvSelectedParameters.ItemsSource = SelectedParameterItems;
            lvSchedules.ItemsSource = ScheduleItems;
        }

        #endregion

        #region Progress Bar Methods

        private void ProgressTimer_Tick(object sender, EventArgs e)
        {
            if (_currentProgress < _targetProgress)
            {
                _currentProgress = Math.Min(_currentProgress + 2, _targetProgress);
                UpdateProgressBarImmediate(_currentProgress);
            }
            else if (_currentProgress > _targetProgress)
            {
                _currentProgress = Math.Max(_currentProgress - 2, _targetProgress);
                UpdateProgressBarImmediate(_currentProgress);
            }
            else
            {
                _progressTimer.Stop();
            }
        }

        private void UpdateProgressBarImmediate(int percentage)
        {
            double containerWidth = progressBarContainer.ActualWidth;
            if (containerWidth <= 0) return;

            var fillWidth = (containerWidth * percentage) / 100.0;
            progressBarFill.Width = fillWidth;

            // Set the corner radius. Right side is rounded only at 100%.
            if (percentage >= 100)
            {
                progressBarFill.CornerRadius = new CornerRadius(12.5);
            }
            else
            {
                progressBarFill.CornerRadius = new CornerRadius(12.5, 0, 0, 12.5);
            }

            progressBarText.Text = $"{percentage}%";
        }

        public void UpdateProgressBar(int percentage)
        {
            _targetProgress = percentage;
            if (!_progressTimer.IsEnabled)
            {
                _progressTimer.Start();
            }
        }

        public void ShowProgressBar()
        {
            _currentProgress = 0;
            _targetProgress = 0;
            progressBarFill.Width = 0;
            progressBarFill.CornerRadius = new CornerRadius(12.5, 0, 0, 12.5);
            progressBarText.Text = "0%";
            UpdateProgressBar(5); // Start with 5% to show activity
        }

        public void HideProgressBar()
        {
            _progressTimer.Stop();
            _currentProgress = 0;
            _targetProgress = 0;
            progressBarFill.Width = 0;
            progressBarFill.CornerRadius = new CornerRadius(12.5, 0, 0, 12.5);
            progressBarText.Text = "Ready";
        }

        #endregion

        #region Window Controls

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !_isMaximized)
            {
                DragMove();
            }
        }

        private void btnMinimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void btnMaximize_Click(object sender, RoutedEventArgs e)
        {
            if (_isMaximized)
            {
                this.WindowState = WindowState.Normal;
                btnMaximize.Content = "🗖";
                _isMaximized = false;
            }
            else
            {
                this.WindowState = WindowState.Maximized;
                btnMaximize.Content = "🗗";
                _isMaximized = true;
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

        #endregion

        #region Main Button Events

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            // Check which tab is active
            if (_mainTabControl != null && _mainTabControl.SelectedIndex == 1)
            {
                // Schedules tab is active
                ExportSchedulesToExcel();
            }
            else
            {
                // Categories tab is active
                ExportToExcel();
            }
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            // Check which tab is active
            if (_mainTabControl != null && _mainTabControl.SelectedIndex == 1)
            {
                // Schedules tab is active
                ImportSchedulesFromExcel();
            }
            else
            {
                // Categories tab is active
                ImportFromExcel();
            }
        }

        #endregion

        #region Schedule Methods

        private void LoadSchedules()
        {
            ScheduleItems.Clear();

            // Add "Select All" item
            ScheduleItems.Add(new ScheduleItem("Select All Schedules", true));

            // Get all schedules from the document
            var schedules = _scheduleManager.GetAllSchedules();

            foreach (var schedule in schedules)
            {
                ScheduleItems.Add(new ScheduleItem(schedule));
            }

            UpdateSelectedSchedulesCount();
        }

        private void UpdateSelectedSchedulesCount()
        {
            var selectedCount = SelectedSchedules.Count;
            if (selectedCount == 0)
            {
                txtSelectedSchedulesCount.Text = "No schedules selected";
            }
            else if (selectedCount == 1)
            {
                txtSelectedSchedulesCount.Text = "1 schedule selected";
            }
            else
            {
                txtSelectedSchedulesCount.Text = $"{selectedCount} schedules selected";
            }
        }

        private void ScheduleCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            if (checkBox == null) return;

            var scheduleItem = checkBox.DataContext as ScheduleItem;
            if (scheduleItem == null) return;

            if (scheduleItem.IsSelectAll)
            {
                // Handle "Select All" checkbox
                bool isChecked = checkBox.IsChecked ?? false;
                foreach (var item in ScheduleItems)
                {
                    if (!item.IsSelectAll)
                    {
                        item.IsSelected = isChecked;
                    }
                }
            }

            UpdateSelectedSchedulesCount();
        }

        private void txtScheduleSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.IsFocused)
            {
                string searchText = textBox.Text.ToLower();
                if (searchText == "search schedules...") return;

                var filteredItems = ScheduleItems.Where(s => s.IsSelectAll || s.ScheduleName.ToLower().Contains(searchText));
                lvSchedules.ItemsSource = new ObservableCollection<ScheduleItem>(filteredItems);
            }
        }

        private void txtScheduleSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.Text == "Search schedules...")
            {
                textBox.Text = "";
            }
        }

        private void txtScheduleSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "Search schedules...";
                lvSchedules.ItemsSource = ScheduleItems;
            }
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
        #region Schedule Export/Import

        private void ExportSchedulesToExcel()
        {
            // Validate selections
            if (!SelectedSchedules.Any())
            {
                frmInfoDialog infoDialog = new frmInfoDialog("Please select at least one schedule.");
                infoDialog.ShowDialog();
                return;
            }

            // Prompt user to save Excel file
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel files|*.xlsx";
            saveDialog.Title = "Save Revit Schedules to Excel";

            string defaultFileName = _doc.Title;
            if (string.IsNullOrEmpty(defaultFileName))
            {
                defaultFileName = "RevitScheduleExport";
            }
            saveDialog.FileName = defaultFileName + "_Schedules.xlsx";

            if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string excelFile = saveDialog.FileName;

            // Show progress bar
            Dispatcher.Invoke(() => ShowProgressBar());

            try
            {
                bool includeHeaders = chkIncludeHeaders.IsChecked ?? true;
                bool includeGrandTotals = chkIncludeGrandTotals.IsChecked ?? true;

                Task.Run(() =>
                {
                    try
                    {
                        _scheduleManager.ExportSchedulesToExcel(
                            SelectedSchedules,
                            excelFile,
                            (progress) => Dispatcher.Invoke(() => UpdateProgressBar(progress)),
                            includeHeaders,
                            includeGrandTotals
                        );

                        Dispatcher.Invoke(() =>
                        {
                            UpdateProgressBar(100);
                            System.Threading.Thread.Sleep(500); // Brief pause at 100%
                            HideProgressBar();
                            frmInfoDialog infoDialog = new frmInfoDialog("Schedules exported successfully");
                            infoDialog.ShowDialog();
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            HideProgressBar();
                            TaskDialog.Show("Error", $"Failed to export schedules:\n{ex.Message}");
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                HideProgressBar();
                TaskDialog.Show("Error", $"Failed to export schedules:\n{ex.Message}");
            }
        }

        private void ImportSchedulesFromExcel()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel files|*.xlsx;*.xls";
            openDialog.Title = "Select Excel File with Schedules to Import";

            if (openDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string excelFile = openDialog.FileName;

            // Show progress bar
            Dispatcher.Invoke(() => ShowProgressBar());

            Task.Run(() =>
            {
                try
                {
                    using (Transaction t = new Transaction(_doc, "Import Schedules from Excel"))
                    {
                        t.Start();

                        var errors = _scheduleManager.ImportSchedulesFromExcel(
                            excelFile,
                            (progress) => Dispatcher.Invoke(() => UpdateProgressBar(progress))
                        );

                        t.Commit();

                        Dispatcher.Invoke(() =>
                        {
                            UpdateProgressBar(100);
                            System.Threading.Thread.Sleep(500); // Brief pause at 100%
                            HideProgressBar();

                            if (errors.Any())
                            {
                                var failForm = new frmImportFailed(errors);
                                failForm.ShowDialog();
                            }
                            else
                            {
                                frmInfoDialog infoDialog = new frmInfoDialog("Schedules imported successfully");
                                infoDialog.ShowDialog();
                            }
                        });
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        HideProgressBar();
                        TaskDialog.Show("Error", $"Failed to import schedules:\n{ex.Message}");
                    });
                }
            });
        }

        #endregion

        #region Category Export/Import

        private void ExportToExcel()
        {
            // Validate selections
            if (!SelectedCategoryNames.Any())
            {
                frmInfoDialog infoDialog = new frmInfoDialog("Please select at least one category.");
                infoDialog.ShowDialog();
                return;
            }

            if (!SelectedParameterNames.Any())
            {
                frmInfoDialog infoDialog = new frmInfoDialog("Please select at least one parameter.");
                infoDialog.ShowDialog();
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

            // Collect all UI data BEFORE entering the background thread
            var selectedCategories = CategoryItems
                .Where(item => item.IsSelected && !item.IsSelectAll)
                .ToList();

            var selectedParameterNames = SelectedParameterNames.ToList();
            bool isEntireModel = IsEntireModelChecked;
            ElementId activeViewId = _doc.ActiveView?.Id;

            // Show progress bar on the UI thread
            ShowProgressBar();

            Task.Run(() =>
            {
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

                    // Create color legend
                    CreateParameterColorLegendThread(colorLegendSheet);

                    // Process each category with progress updates
                    int sheetIndex = 1;
                    int totalCategories = selectedCategories.Count;
                    int totalWork = totalCategories * 100; // Each category contributes 100 units
                    int currentWork = 0;

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

                        // Process elements for this category
                        ProcessCategoryForExportThread(worksheet, categoryItem, selectedParameterNames,
                            isEntireModel, activeViewId, (progress) =>
                            {
                                int categoryProgress = currentWork + progress;
                                int overallProgress = (int)((double)categoryProgress / totalWork * 100);
                                Dispatcher.Invoke(() => UpdateProgressBar(Math.Min(overallProgress, 95)));
                            });

                        currentWork += 100;
                    }

                    // Save the workbook
                    excel.DisplayAlerts = false;
                    workbook.SaveAs(excelFile);
                    excel.DisplayAlerts = true;

                    colorLegendSheet.Activate();

                    // Show 100% progress
                    Dispatcher.Invoke(() => UpdateProgressBar(100));
                    System.Threading.Thread.Sleep(500); // Brief pause at 100%

                    // Show info dialog
                    Dispatcher.Invoke(() =>
                    {
                        HideProgressBar();
                        frmInfoDialog infoDialog = new frmInfoDialog("Sheet exported successfully");
                        infoDialog.ShowDialog();
                    });

                    // Open Excel
                    excel.Visible = true;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (ex.HResult == unchecked((int)0x800A03EC))
                    {
                        Dispatcher.Invoke(() =>
                        {
                            HideProgressBar();
                            frmInfoDialog infoDialog = new frmInfoDialog("The Excel file is already open.\nPlease close it and try again.");
                            infoDialog.ShowDialog();
                        });
                    }
                    else
                    {
                        if (excel != null)
                        {
                            excel.DisplayAlerts = true;
                        }
                        Dispatcher.Invoke(() =>
                        {
                            HideProgressBar();
                            TaskDialog.Show("Error", $"Failed to export parameters:\n{ex.Message}");
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (excel != null)
                    {
                        excel.DisplayAlerts = true;
                    }
                    Dispatcher.Invoke(() =>
                    {
                        HideProgressBar();
                        TaskDialog.Show("Error", $"Failed to export parameters:\n{ex.Message}");
                    });
                }
                finally
                {
                    if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    if (colorLegendSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(colorLegendSheet);
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                }
            });
        }

        // Thread-safe version of ProcessCategoryForExport
        private void ProcessCategoryForExportThread(Excel.Worksheet worksheet, ParaExportCategoryItem categoryItem,
            List<string> selectedParameters, bool isEntireModel, ElementId activeViewId, Action<int> progressCallback)
        {
            // Write headers with multi-line text
            Excel.Range elementIdHeader = (Excel.Range)worksheet.Cells[1, 1];
            elementIdHeader.Value2 = "Element ID";
            elementIdHeader.ColumnWidth = 12;

            for (int i = 0; i < selectedParameters.Count; i++)
            {
                string paramName = selectedParameters[i];
                string paramType = "N/A";
                string paramStorageType = "N/A";

                // Get parameter info
                FilteredElementCollector tempCollector;
                if (isEntireModel)
                {
                    tempCollector = new FilteredElementCollector(_doc);
                }
                else
                {
                    tempCollector = new FilteredElementCollector(_doc, activeViewId);
                }

                tempCollector.OfCategoryId(categoryItem.Category.Id);
                tempCollector.WhereElementIsNotElementType();
                Element tempElement = tempCollector.FirstElement();

                if (tempElement != null)
                {
                    Parameter param = null;
                    bool isTypeParam = false;

                    // Check for parameter (instance and type)
                    param = tempElement.LookupParameter(paramName);
                    if (param != null)
                    {
                        paramType = "Instance Parameter";
                    }
                    else
                    {
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
            headerRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            ((Excel.Range)worksheet.Rows[1]).RowHeight = 45;

            // Get elements
            FilteredElementCollector dataCollector;
            if (isEntireModel)
            {
                dataCollector = new FilteredElementCollector(_doc);
            }
            else
            {
                dataCollector = new FilteredElementCollector(_doc, activeViewId);
            }
            dataCollector.OfCategoryId(categoryItem.Category.Id);
            dataCollector.WhereElementIsNotElementType();
            List<Element> elements = dataCollector.ToList();

            int totalElements = elements.Count;
            int processedElements = 0;
            int row = 2;

            foreach (Element element in elements)
            {
                Excel.Range idCell = (Excel.Range)worksheet.Cells[row, 1];
                idCell.Value2 = element.Id.IntegerValue.ToString();
                idCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                idCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                idCell.Borders.Weight = Excel.XlBorderWeight.xlThin;

                for (int col = 0; col < selectedParameters.Count; col++)
                {
                    string paramName = selectedParameters[col];
                    Excel.Range dataCell = (Excel.Range)worksheet.Cells[row, col + 2];

                    Parameter param = element.LookupParameter(paramName);
                    string value = string.Empty;
                    bool isTypeParam = false;

                    if (param != null)
                    {
                        value = Utils.GetParameterValue(element, paramName);
                    }
                    else
                    {
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
                        if (paramName == "Family" || paramName == "Family and Type" || paramName == "Type")
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                        }
                        else if (param.IsReadOnly)
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                        }
                        else if (isTypeParam)
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                        }
                    }
                    else
                    {
                        dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                    }

                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                }
                row++;
                processedElements++;

                // Update progress
                if (processedElements % 10 == 0 || processedElements == totalElements)
                {
                    int percentage = (int)((double)processedElements / totalElements * 100);
                    progressCallback(percentage);
                }
            }

            worksheet.Columns.AutoFit();
        }

        // Thread-safe version of CreateParameterColorLegend
        private void CreateParameterColorLegendThread(Excel.Worksheet colorLegendSheet)
        {
            // Same content as CreateParameterColorLegend but renamed for clarity
            // [Copy the entire content of CreateParameterColorLegend method here]
            // The method doesn't access UI elements so it's already thread-safe
            CreateParameterColorLegend(colorLegendSheet);
        }

        private void ImportFromExcel()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel files|*.xlsx;*.xls";
            openDialog.Title = "Select Excel File to Import";

            if (openDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string excelFile = openDialog.FileName;

            // Set data for the event handler
            _importEventHandler.SetData(excelFile, _doc, this);

            // Raise the external event
            _importExternalEvent.Raise();
        }

        #endregion

        #region Category Methods

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

            CategoryItems.Add(new ParaExportCategoryItem("Select All Categories", true));

            foreach (Category category in availableCategories.OrderBy(c => c.Name))
            {
                CategoryItems.Add(new ParaExportCategoryItem(category));
            }

            lvCategories.ItemsSource = CategoryItems;

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

            var elementInstances = collector
                .WhereElementIsNotElementType()
                .ToList();

            var categoriesWithElements = elementInstances
                .Where(e => e.Category != null)
                .Select(e => e.Category)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .ToList();

            HashSet<string> excludedCategoryNames = new HashSet<string>
            {
                "Survey Point", "Sun Path", "Project Information", "Project Base Point", "Primary Contours", "Material Assets",
                "Legend Components", "Internal Origin", "Cameras", "HVAC Zones", "Pipe Segments", "Area Based Load Type",
                "Circuit Naming Scheme", "<Sketch>", "Center Line", "Center line", "Lines", "Detail Items",
                "Model Lines", "Detail Lines", "<Room Separation>", "<Area Boundary>", "<Space Separation>",
                "Curtain Panel Tags", "Curtain System Tags", "Detail Item Tags", "Door Tags", "Floor Tags",
                "Generic Annotations", "Keynote Tags", "Material Tags", "Multi-Category Tags", "Parking Tags",
                "Plumbing Fixture Tags", "Property Line Segment Tags", "Property Tags", "Revision Clouds",
                "Room Tags", "Space Tags", "Structural Annotations", "Wall Tags", "Window Tags"
            };

            var modelCategories = categoriesWithElements
                .Where(c => (c.CategoryType == CategoryType.Model || c.Name == "Rooms") &&
                           !excludedCategoryNames.Contains(c.Name) &&
                           !c.Name.ToLower().Contains("line") &&
                           !c.Name.ToLower().Contains("sketch"))
                .ToList();

            // Additional check for Rooms category - ensure there are actual room elements
            var finalCategories = new List<Category>();
            foreach (var category in modelCategories)
            {
                FilteredElementCollector catCollector;
                if (IsEntireModelChecked)
                {
                    catCollector = new FilteredElementCollector(_doc);
                }
                else
                {
                    catCollector = new FilteredElementCollector(_doc, _doc.ActiveView.Id);
                }

                catCollector.OfCategoryId(category.Id);
                catCollector.WhereElementIsNotElementType();

                // Only add category if it has at least one element
                if (catCollector.GetElementCount() > 0)
                {
                    finalCategories.Add(category);
                }
            }

            return finalCategories;
        }
        private void CategoryCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            if (checkBox == null) return;

            var categoryItem = checkBox.DataContext as ParaExportCategoryItem;
            if (categoryItem == null) return;

            if (categoryItem.IsSelectAll)
            {
                // Handle "Select All" checkbox
                bool isChecked = checkBox.IsChecked ?? false;
                foreach (var item in CategoryItems)
                {
                    if (!item.IsSelectAll)
                    {
                        item.IsSelected = isChecked;
                    }
                }
            }
            else
            {
                // Update parameters when a category is selected/deselected
                UpdateAvailableParameters();
            }
        }

        private void UpdateAvailableParameters()
        {
            AvailableParameterItems.Clear();

            var selectedCategories = CategoryItems
                .Where(item => item.IsSelected && !item.IsSelectAll)
                .ToList();

            if (!selectedCategories.Any()) return;

            // Get all unique parameters from selected categories
            HashSet<string> parameterNames = new HashSet<string>();
            Dictionary<string, Parameter> parameterMap = new Dictionary<string, Parameter>();
            Dictionary<string, bool> isReadOnlyMap = new Dictionary<string, bool>();
            Dictionary<string, bool> isTypeParamMap = new Dictionary<string, bool>();

            foreach (var categoryItem in selectedCategories)
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

                collector.OfCategoryId(categoryItem.Category.Id);
                collector.WhereElementIsNotElementType();

                var elements = collector.ToElements();
                if (!elements.Any()) continue;

                foreach (Element element in elements.Take(10)) // Sample first 10 elements
                {
                    // Get instance parameters
                    foreach (Parameter param in element.Parameters)
                    {
                        if (param.Definition != null)
                        {
                            string paramName = param.Definition.Name;
                            if (!parameterNames.Contains(paramName))
                            {
                                parameterNames.Add(paramName);
                                parameterMap[paramName] = param;
                                isReadOnlyMap[paramName] = param.IsReadOnly;
                                isTypeParamMap[paramName] = false;
                            }
                        }
                    }

                    // Get type parameters
                    Element elementType = _doc.GetElement(element.GetTypeId());
                    if (elementType != null)
                    {
                        foreach (Parameter param in elementType.Parameters)
                        {
                            if (param.Definition != null)
                            {
                                string paramName = param.Definition.Name;
                                if (!parameterNames.Contains(paramName))
                                {
                                    parameterNames.Add(paramName);
                                    parameterMap[paramName] = param;
                                    isReadOnlyMap[paramName] = param.IsReadOnly;
                                    isTypeParamMap[paramName] = true;
                                }
                            }
                        }
                    }

                    // Add specific built-in parameters
                    AddSpecificBuiltInParameters(element, parameterNames, parameterMap);
                }
            }

            // Exclude "Type Name" and "Family Name" parameters
            parameterNames.Remove("Type Name");
            parameterNames.Remove("Family Name");

            // Create parameter items
            foreach (string paramName in parameterNames.OrderBy(p => p))
            {
                if (parameterMap.ContainsKey(paramName))
                {
                    bool isReadOnly = isReadOnlyMap.ContainsKey(paramName) && isReadOnlyMap[paramName];
                    bool isTypeParam = isTypeParamMap.ContainsKey(paramName) && isTypeParamMap[paramName];

                    // Make the "Type", "Family and Type", and "Family" parameter read-only
                    if (paramName == "Type" || paramName == "Family and Type" || paramName == "Family")
                    {
                        isReadOnly = true;
                    }

                    AvailableParameterItems.Add(new ParaExportParameterItem(parameterMap[paramName], isReadOnly, isTypeParam));
                }
            }

            // Reset the search text
            txtParameterSearch.Text = "Search parameters...";
        }

        private void AddSpecificBuiltInParameters(Element element, HashSet<string> parameterNames, Dictionary<string, Parameter> parameterMap)
        {
            var builtInParamsToCheck = new Dictionary<string, BuiltInParameter>
            {
                { "Family", BuiltInParameter.ELEM_FAMILY_PARAM },
                { "Family and Type", BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM },
                { "Type", BuiltInParameter.ELEM_TYPE_PARAM },
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

        #endregion

        #region Category Search Methods

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

        #endregion

        #region Parameter Search Methods

        private void txtParameterSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.IsFocused)
            {
                string searchText = textBox.Text.ToLower();

                if (searchText == "search parameters...")
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

        #endregion

        #region Parameter Movement Methods

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

                // Re-apply the current search filter
                string searchText = txtParameterSearch.Text.ToLower();
                if (string.IsNullOrWhiteSpace(searchText) || searchText == "search parameters...")
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

            // Re-apply the current search filter
            string searchText = txtParameterSearch.Text.ToLower();
            if (string.IsNullOrWhiteSpace(searchText) || searchText == "search parameters...")
            {
                lvAvailableParameters.ItemsSource = AvailableParameterItems;
            }
            else
            {
                var filteredParameters = AvailableParameterItems.Where(p => p.ParameterName.ToLower().Contains(searchText));
                lvAvailableParameters.ItemsSource = new ObservableCollection<ParaExportParameterItem>(filteredParameters);
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

            // Re-apply the current search filter
            string searchText = txtParameterSearch.Text.ToLower();
            if (string.IsNullOrWhiteSpace(searchText) || searchText == "search parameters...")
            {
                lvAvailableParameters.ItemsSource = AvailableParameterItems;
            }
            else
            {
                var filteredParameters = AvailableParameterItems.Where(p => p.ParameterName.ToLower().Contains(searchText));
                lvAvailableParameters.ItemsSource = new ObservableCollection<ParaExportParameterItem>(filteredParameters);
            }
        }

        private void btnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();
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
            var selectedItems = lvSelectedParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();
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

        #endregion
        #region Excel Export Helper Methods

        private void ProcessCategoryForExport(Excel.Worksheet worksheet, ParaExportCategoryItem categoryItem, Action<int> progressCallback)
        {
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

                // Get parameter info
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

                if (tempElement != null)
                {
                    Parameter param = null;
                    bool isTypeParam = false;

                    // Check for parameter (instance and type)
                    param = tempElement.LookupParameter(paramName);
                    if (param != null)
                    {
                        paramType = "Instance Parameter";
                    }
                    else
                    {
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
            headerRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            ((Excel.Range)worksheet.Rows[1]).RowHeight = 45;

            // Get elements
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

            int totalElements = elements.Count;
            int processedElements = 0;
            int row = 2;

            foreach (Element element in elements)
            {
                Excel.Range idCell = (Excel.Range)worksheet.Cells[row, 1];
                idCell.Value2 = element.Id.IntegerValue.ToString();
                idCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                idCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                idCell.Borders.Weight = Excel.XlBorderWeight.xlThin;

                for (int col = 0; col < selectedParameters.Count; col++)
                {
                    string paramName = selectedParameters[col];
                    Excel.Range dataCell = (Excel.Range)worksheet.Cells[row, col + 2];

                    Parameter param = element.LookupParameter(paramName);
                    string value = string.Empty;
                    bool isTypeParam = false;

                    if (param != null)
                    {
                        value = Utils.GetParameterValue(element, paramName);
                    }
                    else
                    {
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
                        if (paramName == "Family" || paramName == "Family and Type" || paramName == "Type")
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                        }
                        else if (param.IsReadOnly)
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FF4747"));
                        }
                        else if (isTypeParam)
                        {
                            dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFE699"));
                        }
                    }
                    else
                    {
                        dataCell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#D3D3D3"));
                    }

                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataCell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                }
                row++;
                processedElements++;

                // Update progress
                if (processedElements % 10 == 0 || processedElements == totalElements)
                {
                    int percentage = (int)((double)processedElements / totalElements * 100);
                    progressCallback(percentage);
                }
            }

            worksheet.Columns.AutoFit();
        }

        private void CreateParameterColorLegend(Excel.Worksheet colorLegendSheet)
        {
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
        }

        #endregion
    }

    #region Data Classes

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

        public ParaExportCategoryItem(Category category)
        {
            Category = category;
            CategoryName = category.Name;
            IsSelected = false;
            IsSelectAll = false;
        }

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

        public ParaExportParameterItem(Parameter parameter, bool isReadOnly, bool isTypeParam)
        {
            Parameter = parameter;
            ParameterName = parameter.Definition.Name;
            IsSelected = false;
            IsSelectAll = false;

            if (isReadOnly)
            {
                ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FF4747"));
            }
            else if (isTypeParam)
            {
                ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FFE699"));
            }
            else
            {
                ParameterColor = new SolidColorBrush(Colors.White);
            }
        }

        public ParaExportParameterItem(string displayName, bool isSelectAll = false)
        {
            Parameter = null;
            ParameterName = displayName;
            IsSelected = false;
            IsSelectAll = isSelectAll;
            ParameterColor = new SolidColorBrush(Colors.White);
        }
    }

    #endregion
}