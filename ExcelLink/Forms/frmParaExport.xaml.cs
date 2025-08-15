using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelLink.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

namespace ExcelLink.Forms
{
    public partial class frmParaExport : Window, INotifyPropertyChanged
    {
        private Document _doc;
        private ExternalEvent _importExternalEvent;
        private ImportEventHandler _importEventHandler;
        private ObservableCollection<ParaExportCategoryItem> _categoryItems;
        private ObservableCollection<ParaExportParameterItem> _availableParameterItems;
        private ObservableCollection<ParaExportParameterItem> _selectedParameterItems;
        private ObservableCollection<ScheduleItem> _scheduleItems;
        private ObservableCollection<ParaExportParameterItem> _scheduleParameterItems;

        private ScheduleManager _scheduleManager;
        private bool _isMaximized = false;
        private DispatcherTimer _progressTimer;
        private int _targetProgress;
        private int _currentProgress;
        private System.Windows.Controls.TabControl _mainTabControl;
        private Action _postProgressAction;

        private Excel.Application _exportedExcelApp;
        private Excel.Workbook _exportedExcelWorkbook;

        #region Properties

        public ObservableCollection<ParaExportCategoryItem> CategoryItems
        {
            get { return _categoryItems; }
            set { _categoryItems = value; OnPropertyChanged(nameof(CategoryItems)); }
        }

        public ObservableCollection<ParaExportParameterItem> AvailableParameterItems
        {
            get { return _availableParameterItems; }
            set { _availableParameterItems = value; OnPropertyChanged(nameof(AvailableParameterItems)); }
        }

        public ObservableCollection<ParaExportParameterItem> SelectedParameterItems
        {
            get { return _selectedParameterItems; }
            set { _selectedParameterItems = value; OnPropertyChanged(nameof(SelectedParameterItems)); }
        }

        public ObservableCollection<ScheduleItem> ScheduleItems
        {
            get { return _scheduleItems; }
            set { _scheduleItems = value; OnPropertyChanged(nameof(ScheduleItems)); }
        }

        public ObservableCollection<ParaExportParameterItem> ScheduleParameterItems
        {
            get { return _scheduleParameterItems; }
            set { _scheduleParameterItems = value; OnPropertyChanged(nameof(ScheduleParameterItems)); }
        }

        public bool IsEntireModelChecked => (bool)rbEntireModel.IsChecked;
        public bool IsActiveViewChecked => (bool)rbActiveView.IsChecked;

        public List<string> SelectedCategoryNames => CategoryItems
            .Where(item => item.IsSelected && !item.IsSelectAll)
            .Select(item => item.CategoryName)
            .ToList();

        public List<string> SelectedParameterNames => SelectedParameterItems
            .Select(item => item.ParameterName)
            .ToList();

        public List<ViewSchedule> SelectedSchedules => ScheduleItems
            .Where(item => item.IsSelected && !item.IsSelectAll && item.Schedule != null)
            .Select(item => item.Schedule)
            .ToList();

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

            CategoryItems = new ObservableCollection<ParaExportCategoryItem>();
            AvailableParameterItems = new ObservableCollection<ParaExportParameterItem>();
            SelectedParameterItems = new ObservableCollection<ParaExportParameterItem>();
            ScheduleItems = new ObservableCollection<ScheduleItem>();
            ScheduleParameterItems = new ObservableCollection<ParaExportParameterItem>();

            _progressTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(20)
            };
            _progressTimer.Tick += ProgressTimer_Tick;

            _mainTabControl = this.FindName("mainTabControl") as System.Windows.Controls.TabControl;

            LoadCategoriesBasedOnScope();
            LoadSchedules();

            lvAvailableParameters.ItemsSource = AvailableParameterItems;
            lvSelectedParameters.ItemsSource = SelectedParameterItems;
            lvSchedules.ItemsSource = ScheduleItems;
            lvScheduleParameters.ItemsSource = ScheduleParameterItems;

            // Set initial placeholder states
            txtCategorySearch.Foreground = System.Windows.Media.Brushes.Gray;
            txtParameterSearch.Foreground = System.Windows.Media.Brushes.Gray;
            txtScheduleSearch.Foreground = System.Windows.Media.Brushes.Gray;
        }

        #endregion

        // ... (Progress Bar, Window Controls, Main Button Events methods are unchanged) ...
        #region Progress Bar Methods

        private void ProgressTimer_Tick(object sender, EventArgs e)
        {
            if (_currentProgress < _targetProgress)
            {
                _currentProgress = Math.Min(_currentProgress + 2, _targetProgress);
            }
            else if (_currentProgress > _targetProgress)
            {
                _currentProgress = Math.Max(_currentProgress - 2, _targetProgress);
            }

            UpdateProgressBarImmediate(_currentProgress);

            if (_currentProgress >= 100)
            {
                _progressTimer.Stop();
                _postProgressAction?.Invoke();
                _postProgressAction = null;
            }
        }


        private void UpdateProgressBarImmediate(int percentage)
        {
            double containerWidth = progressBarContainer.ActualWidth;
            if (containerWidth <= 0) return;

            var fillWidth = (containerWidth * percentage) / 100.0;
            progressBarFill.Width = fillWidth;

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
            _targetProgress = Math.Min(100, percentage);
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
            UpdateProgressBar(5);
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
            if (_mainTabControl?.SelectedIndex == 1)
            {
                ExportSchedulesToExcel();
            }
            else
            {
                ExportToExcel();
            }
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            if (_mainTabControl?.SelectedIndex == 1)
            {
                ImportSchedulesFromExcel();
            }
            else
            {
                ImportFromExcel();
            }
        }

        #endregion

        #region Schedule Methods

        private void LoadSchedules()
        {
            ScheduleItems.Clear();
            ScheduleItems.Add(new ScheduleItem("Select All Schedules", true));
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
            txtSelectedSchedulesCount.Text = selectedCount == 1 ? "1 schedule selected" : $"{selectedCount} schedules selected";
        }

        private void ScheduleCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            if (checkBox?.DataContext is ScheduleItem scheduleItem)
            {
                if (scheduleItem.IsSelectAll)
                {
                    bool isChecked = checkBox.IsChecked ?? false;
                    foreach (var item in ScheduleItems.Where(i => !i.IsSelectAll))
                    {
                        item.IsSelected = isChecked;
                    }
                }
                UpdateSelectedSchedulesCount();
                UpdateScheduleParameters();
            }
        }

        private void UpdateScheduleParameters()
        {
            ScheduleParameterItems.Clear();
            var selectedSchedules = SelectedSchedules;

            if (!selectedSchedules.Any()) return;

            var allParameterItems = new List<ParaExportParameterItem>();
            var processedParamNames = new HashSet<string>();

            foreach (var schedule in selectedSchedules)
            {
                var definition = schedule.Definition;
                var sampleElement = new FilteredElementCollector(_doc, schedule.Id).FirstElement();
                Element sampleTypeElement = null;
                if (sampleElement != null)
                {
                    sampleTypeElement = _doc.GetElement(sampleElement.GetTypeId());
                }

                for (int i = 0; i < definition.GetFieldCount(); i++)
                {
                    var field = definition.GetField(i);
                    var paramName = field.GetName();

                    if (processedParamNames.Contains(paramName)) continue;

                    processedParamNames.Add(paramName);

                    Parameter param = null;
                    bool isType = false;
                    bool isReadOnly = false;

                    if (sampleElement != null)
                    {
                        param = _scheduleManager.GetParameterByField(sampleElement, field);
                        if (param == null && sampleTypeElement != null)
                        {
                            param = _scheduleManager.GetParameterByField(sampleTypeElement, field);
                            isType = true;
                        }
                    }

                    if (param != null)
                    {
                        isReadOnly = param.IsReadOnly;
                        allParameterItems.Add(new ParaExportParameterItem(param, isReadOnly, isType));
                    }
                    else
                    {
                        // Handle non-parameter fields like 'Count' or calculated fields
                        allParameterItems.Add(new ParaExportParameterItem(paramName));
                    }
                }
            }

            // Sort and update the UI collection
            var sortedItems = allParameterItems.OrderBy(p => p.ParameterName);
            foreach (var item in sortedItems)
            {
                ScheduleParameterItems.Add(item);
            }
        }


        private void txtScheduleSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            // Only filter if the text is not the placeholder
            if (textBox != null && textBox.IsFocused && textBox.Text != "Search schedules...")
            {
                string searchText = textBox.Text.ToLower();
                var filteredItems = ScheduleItems.Where(s => s.IsSelectAll || s.ScheduleName.ToLower().Contains(searchText));
                lvSchedules.ItemsSource = new ObservableCollection<ScheduleItem>(filteredItems);
            }
        }

        private void txtScheduleSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.Text == "Search schedules...")
            {
                textBox.Text = "";
                textBox.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void txtScheduleSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "Search schedules...";
                textBox.Foreground = System.Windows.Media.Brushes.Gray;
                lvSchedules.ItemsSource = ScheduleItems;
            }
        }

        #endregion

        #region INotifyPropertyChanged
        // ... (unchanged)
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region Schedule Export/Import
        // ... (unchanged)
        private void ExportSchedulesToExcel()
        {
            if (!SelectedSchedules.Any())
            {
                frmInfoDialog infoDialog = new frmInfoDialog("Please select at least one schedule.");
                infoDialog.ShowDialog();
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel files|*.xlsx",
                Title = "Save Revit Schedules to Excel",
                FileName = $"{_doc.Title}_Schedules.xlsx"
            };

            if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            string excelFile = saveDialog.FileName;

            ShowProgressBar();

            var schedulesToExport = SelectedSchedules.ToList();

            List<SimpleScheduleData> scheduleData;
            try
            {
                // =================================================================================
                // MODIFICATION: Changed 'false' to 'true' to include header/column properties.
                // =================================================================================
                scheduleData = _scheduleManager.GetScheduleDataForExport(schedulesToExport, true);
            }
            catch (Exception ex)
            {
                HideProgressBar();
                TaskDialog.Show("Error", $"Failed to read schedule data:\n{ex.Message}");
                return;
            }

            Task.Run(() =>
            {
                try
                {
                    _scheduleManager.ExportSchedulesToExcel(scheduleData, excelFile, (progress) =>
                    {
                        Dispatcher.Invoke(() => UpdateProgressBar(progress));
                    });

                    _postProgressAction = () =>
                    {
                        frmInfoDialog infoDialog = new frmInfoDialog("Schedules exported successfully.");
                        infoDialog.ShowDialog();
                        HideProgressBar();
                        Process.Start(excelFile);
                    };

                    Dispatcher.Invoke(() => UpdateProgressBar(100));
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
                        Dispatcher.Invoke(() =>
                        {
                            HideProgressBar();
                            TaskDialog.Show("Error", $"An Excel error occurred:\n{ex.Message}");
                        });
                    }
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


        public void HandleImportCompletion()
        {
            _postProgressAction = () =>
            {
                var errorMessages = _importEventHandler.ErrorMessages;
                var updatedElementsCount = _importEventHandler.UpdatedElementsCount;

                if (errorMessages.Any())
                {
                    var failForm = new frmImportFailed(errorMessages);
                    failForm.ShowDialog();
                }
                else if (updatedElementsCount > 0)
                {
                    frmInfoDialog infoDialog = new frmInfoDialog("Model updated successfully");
                    infoDialog.ShowDialog();
                }
                else
                {
                    frmInfoDialog infoDialog = new frmInfoDialog("Import completed. \nNo parameters were updated.");
                    infoDialog.ShowDialog();
                }
                HideProgressBar();
            };
            UpdateProgressBar(100);
        }

        private void ImportSchedulesFromExcel()
        {
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Filter = "Excel files|*.xlsx;*.xls",
                Title = "Select Excel File with Schedules to Import"
            };

            if (openDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            string excelFile = openDialog.FileName;

            ShowProgressBar();

            Task.Run(() =>
            {
                try
                {
                    List<ImportErrorItem> errors;
                    using (Transaction t = new Transaction(_doc, "Import Schedules from Excel"))
                    {
                        t.Start();
                        errors = _scheduleManager.ImportSchedulesFromExcel(
                            excelFile,
                            (progress) => Dispatcher.Invoke(() => UpdateProgressBar(progress))
                        );
                        t.Commit();
                    }

                    _postProgressAction = () =>
                    {
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
                        HideProgressBar();
                    };

                    Dispatcher.Invoke(() => UpdateProgressBar(100));
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

        #region Category and Parameter Logic

        private void txtCategorySearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.IsFocused && textBox.Text != "Search categories...")
            {
                string searchText = textBox.Text.ToLower();
                var filteredItems = CategoryItems.Where(c => c.IsSelectAll || c.CategoryName.ToLower().Contains(searchText));
                lvCategories.ItemsSource = new ObservableCollection<ParaExportCategoryItem>(filteredItems);
            }
        }

        private void txtCategorySearch_GotFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.Text == "Search categories...")
            {
                textBox.Text = "";
                textBox.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void txtCategorySearch_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "Search categories...";
                textBox.Foreground = System.Windows.Media.Brushes.Gray;
                lvCategories.ItemsSource = CategoryItems;
            }
        }

        private void txtParameterSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if ((sender as System.Windows.Controls.TextBox)?.IsFocused == true && (sender as System.Windows.Controls.TextBox).Text != "Search parameters...")
            {
                ApplyParameterFilter();
            }
        }

        private void txtParameterSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.Text == "Search parameters...")
            {
                textBox.Text = "";
                textBox.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void txtParameterSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            var textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = "Search parameters...";
                textBox.Foreground = System.Windows.Media.Brushes.Gray;
                ApplyParameterFilter();
            }
        }

        // ... (The rest of the file is unchanged)
        private void ExportToExcel()
        {
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

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel files|*.xlsx",
                Title = "Save Revit Parameters to Excel",
                FileName = $"{_doc.Title}.xlsx"
            };

            if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            string excelFile = saveDialog.FileName;
            var selectedCategories = CategoryItems.Where(item => item.IsSelected && !item.IsSelectAll).ToList();
            var selectedParameterNames = SelectedParameterNames.ToList();
            bool isEntireModel = IsEntireModelChecked;
            ElementId activeViewId = _doc.ActiveView?.Id;

            ShowProgressBar();

            Task.Run(() =>
            {
                bool exportSuccess = false;
                try
                {
                    _exportedExcelApp = new Excel.Application();
                    _exportedExcelWorkbook = _exportedExcelApp.Workbooks.Add();

                    while (_exportedExcelWorkbook.Worksheets.Count > 1)
                    {
                        ((Excel.Worksheet)_exportedExcelWorkbook.Worksheets[_exportedExcelWorkbook.Worksheets.Count]).Delete();
                    }

                    Excel.Worksheet colorLegendSheet = (Excel.Worksheet)_exportedExcelWorkbook.Worksheets[1];
                    colorLegendSheet.Name = "Color Legend";
                    CreateParameterColorLegend(colorLegendSheet);

                    int sheetIndex = 1;
                    int totalCategories = selectedCategories.Count;
                    int totalWork = totalCategories * 100;
                    int currentWork = 0;

                    foreach (var categoryItem in selectedCategories)
                    {
                        sheetIndex++;
                        Excel.Worksheet worksheet;
                        if (_exportedExcelWorkbook.Worksheets.Count < sheetIndex)
                        {
                            worksheet = (Excel.Worksheet)_exportedExcelWorkbook.Worksheets.Add(After: _exportedExcelWorkbook.Worksheets[_exportedExcelWorkbook.Worksheets.Count]);
                        }
                        else
                        {
                            worksheet = (Excel.Worksheet)_exportedExcelWorkbook.Worksheets[sheetIndex];
                        }

                        string sheetName = categoryItem.CategoryName.Length > 31 ? categoryItem.CategoryName.Substring(0, 31) : categoryItem.CategoryName;
                        worksheet.Name = sheetName;

                        ProcessCategoryForExport(worksheet, categoryItem, selectedParameterNames, isEntireModel, activeViewId, (progress) =>
                        {
                            int categoryProgress = currentWork + progress;
                            int overallProgress = (int)((double)categoryProgress / totalWork * 100);
                            Dispatcher.Invoke(() => UpdateProgressBar(Math.Min(overallProgress, 95)));
                        });

                        currentWork += 100;
                    }

                    _exportedExcelApp.DisplayAlerts = false;
                    _exportedExcelWorkbook.SaveAs(excelFile);
                    _exportedExcelApp.DisplayAlerts = true;
                    colorLegendSheet.Activate();

                    _postProgressAction = () => {
                        frmInfoDialog infoDialog = new frmInfoDialog("Sheet exported successfully");
                        infoDialog.ShowDialog();
                        if (_exportedExcelApp != null)
                        {
                            _exportedExcelApp.Visible = true;
                        }
                        HideProgressBar();
                    };

                    Dispatcher.Invoke(() => UpdateProgressBar(100));
                    exportSuccess = true;
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
                        if (_exportedExcelApp != null) _exportedExcelApp.DisplayAlerts = true;
                        Dispatcher.Invoke(() =>
                        {
                            HideProgressBar();
                            TaskDialog.Show("Error", $"An Excel error occurred:\n{ex.Message}");
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (_exportedExcelApp != null) _exportedExcelApp.DisplayAlerts = true;
                    Dispatcher.Invoke(() =>
                    {
                        HideProgressBar();
                        TaskDialog.Show("Error", $"Failed to export parameters:\n{ex.Message}");
                    });
                }
                finally
                {
                    if (!exportSuccess)
                    {
                        if (_exportedExcelWorkbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(_exportedExcelWorkbook);
                        if (_exportedExcelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(_exportedExcelApp);
                    }
                }
            });
        }

        private void ImportFromExcel()
        {
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Filter = "Excel files|*.xlsx;*.xls",
                Title = "Select Excel File to Import"
            };

            if (openDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            string excelFile = openDialog.FileName;

            ShowProgressBar();
            _importEventHandler.SetData(excelFile, _doc, this);
            _importExternalEvent.Raise();
        }

        private void rbEntireModel_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) LoadCategoriesBasedOnScope();
        }

        private void rbActiveView_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) LoadCategoriesBasedOnScope();
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
            FilteredElementCollector collector = IsEntireModelChecked ? new FilteredElementCollector(_doc) : new FilteredElementCollector(_doc, _doc.ActiveView.Id);

            var elementInstances = collector.WhereElementIsNotElementType().ToList();
            var categoriesWithElements = elementInstances
                .Where(e => e.Category != null)
                .Select(e => e.Category)
                .GroupBy(c => c.Id)
                .Select(g => g.First())
                .ToList();

            HashSet<string> excludedCategoryNames = new HashSet<string> { "Survey Point", "Sun Path", "Project Information", "Project Base Point", "Primary Contours", "Material Assets", "Legend Components", "Internal Origin", "Cameras", "HVAC Zones", "Pipe Segments", "Area Based Load Type", "Circuit Naming Scheme", "<Sketch>", "Center Line", "Center line", "Lines", "Detail Items", "Model Lines", "Detail Lines", "<Room Separation>", "<Area Boundary>", "<Space Separation>", "Curtain Panel Tags", "Curtain System Tags", "Detail Item Tags", "Door Tags", "Floor Tags", "Generic Annotations", "Keynote Tags", "Material Tags", "Multi-Category Tags", "Parking Tags", "Plumbing Fixture Tags", "Property Line Segment Tags", "Property Tags", "Revision Clouds", "Room Tags", "Space Tags", "Structural Annotations", "Wall Tags", "Window Tags" };

            var modelCategories = categoriesWithElements
                .Where(c => (c.CategoryType == CategoryType.Model || c.Name == "Rooms") && !excludedCategoryNames.Contains(c.Name) && !c.Name.ToLower().Contains("line") && !c.Name.ToLower().Contains("sketch"))
                .ToList();

            var finalCategories = new List<Category>();
            foreach (var category in modelCategories)
            {
                FilteredElementCollector catCollector = IsEntireModelChecked ? new FilteredElementCollector(_doc) : new FilteredElementCollector(_doc, _doc.ActiveView.Id);
                catCollector.OfCategoryId(category.Id).WhereElementIsNotElementType();
                if (catCollector.Any())
                {
                    finalCategories.Add(category);
                }
            }
            return finalCategories;
        }

        private void CategoryCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            if (checkBox?.DataContext is ParaExportCategoryItem categoryItem)
            {
                if (categoryItem.IsSelectAll)
                {
                    bool isChecked = checkBox.IsChecked ?? false;
                    foreach (var item in CategoryItems.Where(i => !i.IsSelectAll))
                    {
                        item.IsSelected = isChecked;
                    }
                }
                UpdateAvailableParameters();
            }
        }

        private void UpdateAvailableParameters()
        {
            AvailableParameterItems.Clear();
            var selectedCategories = CategoryItems.Where(item => item.IsSelected && !item.IsSelectAll).ToList();
            if (!selectedCategories.Any()) return;

            var parameterNames = new HashSet<string>();
            var parameterMap = new Dictionary<string, Parameter>();
            var isReadOnlyMap = new Dictionary<string, bool>();
            var isTypeParamMap = new Dictionary<string, bool>();

            foreach (var categoryItem in selectedCategories)
            {
                FilteredElementCollector collector = IsEntireModelChecked ? new FilteredElementCollector(_doc) : new FilteredElementCollector(_doc, _doc.ActiveView.Id);
                collector.OfCategoryId(categoryItem.Category.Id).WhereElementIsNotElementType();
                var elements = collector.ToElements();
                if (!elements.Any()) continue;

                foreach (Element element in elements.Take(10))
                {
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
                    AddSpecificBuiltInParameters(element, parameterNames, parameterMap);
                }
            }

            parameterNames.ExceptWith(new[] { "Type Name", "Family Name", "Category", "Type Id" });

            foreach (string paramName in parameterNames.OrderBy(p => p))
            {
                if (parameterMap.ContainsKey(paramName))
                {
                    bool isReadOnly = isReadOnlyMap.ContainsKey(paramName) ? isReadOnlyMap[paramName] : false;
                    bool isTypeParam = isTypeParamMap.ContainsKey(paramName) ? isTypeParamMap[paramName] : false;

                    if (paramName == "Type" || paramName == "Family and Type" || paramName == "Family") isReadOnly = true;
                    AvailableParameterItems.Add(new ParaExportParameterItem(parameterMap[paramName], isReadOnly, isTypeParam));
                }
            }
            txtParameterSearch.Text = "Search parameters...";
        }

        private void AddSpecificBuiltInParameters(Element element, HashSet<string> parameterNames, Dictionary<string, Parameter> parameterMap)
        {
            var builtInParamsToCheck = new Dictionary<string, BuiltInParameter> { { "Family", BuiltInParameter.ELEM_FAMILY_PARAM }, { "Family and Type", BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM }, { "Type", BuiltInParameter.ELEM_TYPE_PARAM }, { "Comments", BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS }, { "Type Comments", BuiltInParameter.ALL_MODEL_TYPE_COMMENTS }, { "Mark", BuiltInParameter.ALL_MODEL_MARK }, { "Type Mark", BuiltInParameter.ALL_MODEL_TYPE_MARK }, { "Description", BuiltInParameter.ALL_MODEL_DESCRIPTION }, { "Manufacturer", BuiltInParameter.ALL_MODEL_MANUFACTURER }, { "Model", BuiltInParameter.ALL_MODEL_MODEL }, { "URL", BuiltInParameter.ALL_MODEL_URL }, { "Cost", BuiltInParameter.ALL_MODEL_COST }, { "Assembly Code", BuiltInParameter.UNIFORMAT_CODE }, { "Assembly Description", BuiltInParameter.UNIFORMAT_DESCRIPTION }, { "Keynote", BuiltInParameter.KEYNOTE_PARAM }, { "Area", BuiltInParameter.HOST_AREA_COMPUTED }, { "Volume", BuiltInParameter.HOST_VOLUME_COMPUTED }, { "Perimeter", BuiltInParameter.HOST_PERIMETER_COMPUTED }, { "Level", BuiltInParameter.LEVEL_PARAM } };

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
                            if (elementType != null) param = elementType.get_Parameter(kvp.Value);
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

        private void ApplyParameterFilter()
        {
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

        private void lvAvailableParameters_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lvAvailableParameters.SelectedItem is ParaExportParameterItem item)
            {
                MoveParameters(new[] { item }, AvailableParameterItems, SelectedParameterItems);
            }
        }

        private void lvSelectedParameters_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (lvSelectedParameters.SelectedItem is ParaExportParameterItem item)
            {
                MoveParameters(new[] { item }, SelectedParameterItems, AvailableParameterItems);
            }
        }

        private void MoveParameters(IEnumerable<ParaExportParameterItem> itemsToMove, ObservableCollection<ParaExportParameterItem> source, ObservableCollection<ParaExportParameterItem> destination)
        {
            var movedItems = itemsToMove.ToList();
            foreach (var item in movedItems)
            {
                if (source.Remove(item)) destination.Add(item);
            }
            if (destination == AvailableParameterItems)
            {
                var sorted = destination.OrderBy(p => p.ParameterName).ToList();
                destination.Clear();
                foreach (var item in sorted) destination.Add(item);
            }
            ApplyParameterFilter();
        }

        private void btnMoveRight_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvAvailableParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();
            MoveParameters(selectedItems, AvailableParameterItems, SelectedParameterItems);
        }

        private void btnMoveLeft_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();
            MoveParameters(selectedItems, SelectedParameterItems, AvailableParameterItems);
        }

        private void btnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();
            if (!selectedItems.Any()) return;
            foreach (var item in selectedItems)
            {
                int index = SelectedParameterItems.IndexOf(item);
                if (index > 0) SelectedParameterItems.Move(index, index - 1);
            }
        }

        private void btnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = lvSelectedParameters.SelectedItems.Cast<ParaExportParameterItem>().ToList();
            if (!selectedItems.Any()) return;
            for (int i = selectedItems.Count - 1; i >= 0; i--)
            {
                var item = selectedItems[i];
                int index = SelectedParameterItems.IndexOf(item);
                if (index < SelectedParameterItems.Count - 1) SelectedParameterItems.Move(index, index + 1);
            }
        }

        private void btnMoveAllRight_Click(object sender, RoutedEventArgs e)
        {
            string searchText = txtParameterSearch.Text;
            bool isSearching = !string.IsNullOrWhiteSpace(searchText) && searchText != "Search parameters...";
            var itemsToMove = isSearching ? lvAvailableParameters.Items.Cast<ParaExportParameterItem>().ToList() : AvailableParameterItems.ToList();
            MoveParameters(itemsToMove, AvailableParameterItems, SelectedParameterItems);
        }

        private void btnMoveAllLeft_Click(object sender, RoutedEventArgs e)
        {
            var allSelectedItems = SelectedParameterItems.ToList();
            MoveParameters(allSelectedItems, SelectedParameterItems, AvailableParameterItems);
        }

        private void ProcessCategoryForExport(Excel.Worksheet worksheet, ParaExportCategoryItem categoryItem, List<string> selectedParameters, bool isEntireModel, ElementId activeViewId, Action<int> progressCallback)
        {
            Excel.Range elementIdHeader = (Excel.Range)worksheet.Cells[1, 1];
            elementIdHeader.Value2 = "Element ID";
            elementIdHeader.ColumnWidth = 12;

            for (int i = 0; i < selectedParameters.Count; i++)
            {
                string paramName = selectedParameters[i];
                string paramType = "N/A";
                string paramStorageType = "N/A";

                FilteredElementCollector tempCollector = isEntireModel ? new FilteredElementCollector(_doc) : new FilteredElementCollector(_doc, activeViewId);
                tempCollector.OfCategoryId(categoryItem.Category.Id).WhereElementIsNotElementType();
                Element tempElement = tempCollector.FirstElement();

                if (tempElement != null)
                {
                    Parameter param = tempElement.LookupParameter(paramName);
                    bool isTypeParam = false;
                    if (param == null)
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
                    else
                    {
                        paramType = "Instance Parameter";
                    }

                    if (param == null)
                    {
                        BuiltInParameter bip = Utils.GetBuiltInParameterByName(paramName);
                        if (bip != BuiltInParameter.INVALID)
                        {
                            param = tempElement.get_Parameter(bip);
                            if (param == null)
                            {
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
                            else
                            {
                                paramType = "Instance Parameter";
                            }
                        }
                    }
                    if (param != null) paramStorageType = Utils.GetParameterStorageTypeString(param.StorageType);
                }

                string headerText = $"{paramName}{Environment.NewLine}({paramType}){Environment.NewLine}Type: {paramStorageType}";
                Excel.Range headerCell = (Excel.Range)worksheet.Cells[1, i + 2];
                headerCell.Value2 = headerText;
                headerCell.ColumnWidth = Math.Max(15, Math.Min(30, paramName.Length + 5));
            }

            Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, selectedParameters.Count + 1]];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC729"));
            headerRange.WrapText = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            headerRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            headerRange.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, true);
            ((Excel.Range)worksheet.Rows[1]).RowHeight = 45;

            FilteredElementCollector dataCollector = isEntireModel ? new FilteredElementCollector(_doc) : new FilteredElementCollector(_doc, activeViewId);
            dataCollector.OfCategoryId(categoryItem.Category.Id).WhereElementIsNotElementType();
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
                    string value = string.Empty;
                    bool isTypeParam = false;
                    Parameter param = element.LookupParameter(paramName);
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
                        if (paramName == "Family" || paramName == "Family and Type" || paramName == "Type" || param.IsReadOnly)
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
                if (processedElements % 10 == 0 || processedElements == totalElements)
                {
                    progressCallback((int)((double)processedElements / totalElements * 100));
                }
            }
            worksheet.Columns.AutoFit();
        }

        private void CreateParameterColorLegend(Excel.Worksheet colorLegendSheet)
        {
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

            Excel.Range dataRange = colorLegendSheet.Range[colorLegendSheet.Cells[4, 2], colorLegendSheet.Cells[8, 4]];
            dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range entireTable = colorLegendSheet.Range[colorLegendSheet.Cells[3, 2], colorLegendSheet.Cells[8, 4]];
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            entireTable.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            ((Excel.Range)colorLegendSheet.Columns[2]).ColumnWidth = 15;
            ((Excel.Range)colorLegendSheet.Columns[3]).ColumnWidth = 35;
            ((Excel.Range)colorLegendSheet.Columns[4]).ColumnWidth = 50;
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

        public Category Category { get => _category; set { _category = value; OnPropertyChanged(nameof(Category)); } }
        public string CategoryName { get => _categoryName; set { _categoryName = value; OnPropertyChanged(nameof(CategoryName)); } }
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }
        public bool IsSelectAll { get => _isSelectAll; set { _isSelectAll = value; OnPropertyChanged(nameof(IsSelectAll)); OnPropertyChanged(nameof(FontWeight)); OnPropertyChanged(nameof(TextColor)); } }
        public string FontWeight => IsSelectAll ? "Bold" : "Normal";
        public string TextColor => IsSelectAll ? "#000000" : "#000000";

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public ParaExportCategoryItem(Category category)
        {
            Category = category;
            CategoryName = category.Name;
        }
        public ParaExportCategoryItem(string displayName, bool isSelectAll = false)
        {
            CategoryName = displayName;
            IsSelectAll = isSelectAll;
        }
    }

    public class ParaExportParameterItem : INotifyPropertyChanged
    {
        private Parameter _parameter;
        private string _parameterName;
        private SolidColorBrush _parameterColor;

        public Parameter Parameter { get => _parameter; set { _parameter = value; OnPropertyChanged(nameof(Parameter)); } }
        public string ParameterName { get => _parameterName; set { _parameterName = value; OnPropertyChanged(nameof(ParameterName)); } }
        public SolidColorBrush ParameterColor { get => _parameterColor; set { _parameterColor = value; OnPropertyChanged(nameof(ParameterColor)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public ParaExportParameterItem(Parameter parameter, bool isReadOnly, bool isTypeParam)
        {
            Parameter = parameter;
            ParameterName = parameter.Definition.Name;
            if (isReadOnly)
                ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FF4747"));
            else if (isTypeParam)
                ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FFE699"));
            else
                ParameterColor = new SolidColorBrush(Colors.White);
        }

        // Constructor for non-parameter fields like "Count"
        public ParaExportParameterItem(string parameterName)
        {
            Parameter = null;
            ParameterName = parameterName;
            // Treat as read-only
            ParameterColor = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#80FF4747"));
        }
    }

    #endregion
}