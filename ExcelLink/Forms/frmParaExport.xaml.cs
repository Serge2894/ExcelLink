using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelLink.Forms
{
    /// <summary>
    /// Interaction logic for frmParaExport.xaml
    /// </summary>
    public partial class frmParaExport : Window, INotifyPropertyChanged
    {
        private Document _doc;
        private ObservableCollection<ParaExportCategoryItem> _categoryItems;
        private ObservableCollection<ParaExportParameterItem> _parameterItems;

        public ObservableCollection<ParaExportCategoryItem> CategoryItems
        {
            get { return _categoryItems; }
            set
            {
                _categoryItems = value;
                OnPropertyChanged(nameof(CategoryItems));
            }
        }

        public ObservableCollection<ParaExportParameterItem> ParameterItems
        {
            get { return _parameterItems; }
            set
            {
                _parameterItems = value;
                OnPropertyChanged(nameof(ParameterItems));
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
                return ParameterItems
                    .Where(item => item.IsSelected && !item.IsSelectAll)
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
            ParameterItems = new ObservableCollection<ParaExportParameterItem>();

            // Load initial data
            LoadCategoriesBasedOnScope();
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
            DialogResult = false;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            btnExport.Tag = true;
            DialogResult = true;
            Close();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            btnExport.Tag = false;
            DialogResult = true;
            Close();
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
            ParameterItems.Clear();

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
                "Rooms",
                "HVAC Zones",
                "Pipe Segments",
                "Area Based Load Type",
                "Circuit Naming Scheme",
                "<Sketch>" // Excludes the "<Sketch>" category
            };

            // Filter to only include model categories
            var modelCategories = categoriesWithElements
                .Where(c => c.CategoryType == CategoryType.Model &&
                           !excludedCategoryNames.Contains(c.Name))
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
            ParameterItems.Clear();

            var selectedCategories = CategoryItems
                .Where(item => item.IsSelected && !item.IsSelectAll)
                .ToList();

            if (!selectedCategories.Any()) return;

            List<Parameter> allParameters = new List<Parameter>();

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

                    // Collect parameters from instances
                    foreach (Element instance in instances)
                    {
                        // Get instance parameters
                        allParameters.AddRange(GetAllParametersFromElement(instance));

                        // Get type parameters
                        ElementId typeId = instance.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            Element elementType = _doc.GetElement(typeId);
                            if (elementType != null)
                            {
                                allParameters.AddRange(GetAllParametersFromElement(elementType));
                            }
                        }

                        break; // We only need one instance per category to get all parameters
                    }
                }
            }

            // List of parameter names to exclude
            HashSet<string> excludedParameters = new HashSet<string>
            {
                "Family and Type",
                "Family",
                "Type",
                "Phase Created",
                "Phase Demolished",
            };

            // Filter and deduplicate parameters
            var distinctParameters = allParameters
                .Where(p => !p.IsReadOnly &&
                           !excludedParameters.Contains(p.Definition.Name) &&
                           (p.StorageType == StorageType.String ||
                            p.StorageType == StorageType.Double ||
                            p.StorageType == StorageType.Integer))
                .GroupBy(x => x.Definition.Name)
                .Select(x => x.First())
                .OrderBy(x => x.Definition.Name)
                .ToList();

            // Add "Select All" option
            ParameterItems.Add(new ParaExportParameterItem("Select All Parameters", true));

            // Add individual parameters
            foreach (Parameter param in distinctParameters)
            {
                ParameterItems.Add(new ParaExportParameterItem(param));
            }

            // Set ListView source
            lvParameters.ItemsSource = ParameterItems;

            // Initialize search box
            txtParameterSearch.Text = "Search parameters...";
        }

        private List<Parameter> GetAllParametersFromElement(Element element)
        {
            List<Parameter> parameters = new List<Parameter>();

            foreach (Parameter param in element.Parameters)
            {
                if (param != null && param.Definition != null)
                {
                    parameters.Add(param);
                }
            }

            return parameters;
        }

        private void ParameterCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox checkBox && checkBox.DataContext is ParaExportParameterItem paramItem)
            {
                if (paramItem.IsSelectAll)
                {
                    // Handle "Select All" checkbox
                    bool isChecked = checkBox.IsChecked == true;
                    foreach (ParaExportParameterItem item in ParameterItems)
                    {
                        if (!item.IsSelectAll)
                        {
                            item.IsSelected = isChecked;
                        }
                    }
                }
                else
                {
                    // Handle individual parameter checkbox
                    UpdateParameterSelectAllCheckboxState();
                }

                UpdateParameterSearchTextBox();
            }
        }

        private void UpdateParameterSelectAllCheckboxState()
        {
            var selectAllItem = ParameterItems.FirstOrDefault(item => item.IsSelectAll);
            if (selectAllItem != null)
            {
                var parameterItems = ParameterItems.Where(item => !item.IsSelectAll).ToList();
                int selectedCount = parameterItems.Count(item => item.IsSelected);
                int totalCount = parameterItems.Count;

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

        private void UpdateParameterSearchTextBox()
        {
            var selectedItems = ParameterItems.Where(item => item.IsSelected && !item.IsSelectAll).ToList();

            if (selectedItems.Count == 0)
            {
                txtParameterSearch.Text = "Search parameters...";
            }
            else if (selectedItems.Count == 1)
            {
                txtParameterSearch.Text = selectedItems.First().ParameterName;
            }
            else
            {
                txtParameterSearch.Text = $"{selectedItems.Count} parameters selected";
            }
        }

        // Search functionality for categories
        private void txtCategorySearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            if (textBox != null && textBox.IsFocused)
            {
                string searchText = textBox.Text.ToLower();

                if (searchText == "search categories...")
                    return;

                List<ParaExportCategoryItem> filteredItems;

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    filteredItems = CategoryItems.ToList();
                }
                else
                {
                    filteredItems = CategoryItems
                        .Where(c => c.IsSelectAll || c.CategoryName.ToLower().Contains(searchText))
                        .ToList();
                }

                ObservableCollection<ParaExportCategoryItem> filteredCollection = new ObservableCollection<ParaExportCategoryItem>();
                foreach (var item in filteredItems)
                {
                    filteredCollection.Add(item);
                }

                lvCategories.ItemsSource = filteredCollection;
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
                UpdateCategorySearchTextBox();
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

                if (searchText == "search parameters...")
                    return;

                List<ParaExportParameterItem> filteredItems;

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    filteredItems = ParameterItems.ToList();
                }
                else
                {
                    filteredItems = ParameterItems
                        .Where(p => p.IsSelectAll || p.ParameterName.ToLower().Contains(searchText))
                        .ToList();
                }

                ObservableCollection<ParaExportParameterItem> filteredCollection = new ObservableCollection<ParaExportParameterItem>();
                foreach (var item in filteredItems)
                {
                    filteredCollection.Add(item);
                }

                lvParameters.ItemsSource = filteredCollection;
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
                UpdateParameterSearchTextBox();
                lvParameters.ItemsSource = ParameterItems;
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
        public ParaExportParameterItem(Parameter parameter)
        {
            Parameter = parameter;
            ParameterName = parameter.Definition.Name;
            IsSelected = false;
            IsSelectAll = false;
        }

        // Constructor for "Select All" item
        public ParaExportParameterItem(string displayName, bool isSelectAll = false)
        {
            Parameter = null;
            ParameterName = displayName;
            IsSelected = false;
            IsSelectAll = isSelectAll;
        }
    }
}