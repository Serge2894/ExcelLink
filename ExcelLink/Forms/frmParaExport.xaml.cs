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
            ParameterItems.Clear();

            var selectedCategories = CategoryItems
                .Where(item => item.IsSelected && !item.IsSelectAll)
                .ToList();

            if (!selectedCategories.Any()) return;

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

                        break; // We only need one instance per category to get all parameters
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
                               allowedElementIdParameters.Contains(kvp.Key))) && // Only allow specific ElementId parameters
                             kvp.Value.StorageType != StorageType.None)
                .OrderBy(kvp => kvp.Key)
                .ToList();

            // Add "Select All" option
            ParameterItems.Add(new ParaExportParameterItem("Select All Parameters", true));

            // Add individual parameters
            foreach (var kvp in distinctParameters)
            {
                ParameterItems.Add(new ParaExportParameterItem(kvp.Value));
            }

            // Set ListView source
            lvParameters.ItemsSource = ParameterItems;

            // Initialize search box
            txtParameterSearch.Text = "Search parameters...";
        }

        private void AddSpecificBuiltInParameters(Element element, HashSet<string> parameterNames, Dictionary<string, Parameter> parameterMap)
        {
            // List of specific built-in parameters to check
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

            // Add category-specific parameters
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

                // Common geometric parameters
                builtInParamsToCheck["Area"] = BuiltInParameter.HOST_AREA_COMPUTED;
                builtInParamsToCheck["Volume"] = BuiltInParameter.HOST_VOLUME_COMPUTED;
                builtInParamsToCheck["Perimeter"] = BuiltInParameter.HOST_PERIMETER_COMPUTED;
                builtInParamsToCheck["Level"] = BuiltInParameter.LEVEL_PARAM;
            }

            // Try to get each built-in parameter
            foreach (var kvp in builtInParamsToCheck)
            {
                try
                {
                    Parameter param = element.get_Parameter(kvp.Value);
                    if (param == null)
                    {
                        // Try on type
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
                catch
                {
                    // Skip if parameter doesn't exist for this element type
                }
            }
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