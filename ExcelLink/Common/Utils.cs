using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using ExcelLink.Forms;
using System.Diagnostics;
using System.Globalization;

namespace ExcelLink.Common
{
    internal static class Utils
    {
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        // This method finds a parameter by name, first by LookupParameter, then by BuiltInParameter.
        internal static Parameter GetParameterByName(Element curElem, string paramName)
        {
            // First, try to get the parameter directly by name
            Parameter param = curElem.LookupParameter(paramName);
            if (param != null)
            {
                return param;
            }

            // If not found, try to find a matching built-in parameter
            BuiltInParameter bip = GetBuiltInParameterByName(paramName);
            if (bip != BuiltInParameter.INVALID)
            {
                param = curElem.get_Parameter(bip);
            }

            return param;
        }

        // This method gets the parameter value as a formatted string.
        internal static string GetParameterValue(Element curElem, string paramName)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);
            if (curParam != null)
            {
                switch (curParam.StorageType)
                {
                    case StorageType.Double:
                        double dValue = curParam.AsDouble();

#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025
                        // For Revit 2021 and later, use ForgeTypeId
                        try
                        {
                            ForgeTypeId unitTypeId = curParam.GetUnitTypeId();
                            if (unitTypeId != null && !unitTypeId.Empty())
                            {
                                dValue = UnitUtils.ConvertFromInternalUnits(dValue, unitTypeId);
                            }
                        }
                        catch { }
#else
                        // For Revit 2020 and earlier, use DisplayUnitType
                        try
                        {
                            DisplayUnitType dut = curParam.DisplayUnitType;
                            if (dut != DisplayUnitType.DUT_UNDEFINED)
                            {
                                dValue = UnitUtils.ConvertFromInternalUnits(dValue, dut);
                            }
                        }
                        catch { }
#endif
                        return dValue.ToString();

                    case StorageType.Integer:
                        // Check if it's a Yes/No parameter
                        int iValue = curParam.AsInteger();
                        if (IsYesNoParameter(curParam))
                        {
                            return iValue == 1 ? "Yes" : "No";
                        }
                        return iValue.ToString();

                    case StorageType.String:
                        return curParam.AsString() ?? string.Empty;

                    case StorageType.ElementId:
                        ElementId id = curParam.AsElementId();
                        if (id.IntegerValue < 0)
                        {
                            // Built-in category or invalid
                            return string.Empty;
                        }
                        Element elem = curElem.Document.GetElement(id);
                        if (elem != null)
                        {
                            return elem.Name;
                        }
                        return id.IntegerValue.ToString();

                    default:
                        return string.Empty;
                }
            }
            return string.Empty;
        }

        // This method sets the parameter value based on its storage type.
        internal static bool SetParameterValue(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);
            if (curParam != null && !curParam.IsReadOnly)
            {
                try
                {
                    switch (curParam.StorageType)
                    {
                        case StorageType.Double:
                            double dValue;
                            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out dValue))
                            {
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024 || REVIT2025
                                // For Revit 2021 and later, use ForgeTypeId
                                try
                                {
                                    ForgeTypeId unitTypeId = curParam.GetUnitTypeId();
                                    if (unitTypeId != null && !unitTypeId.Empty())
                                    {
                                        dValue = UnitUtils.ConvertToInternalUnits(dValue, unitTypeId);
                                    }
                                }
                                catch { }
#else
                                // For Revit 2020 and earlier, use DisplayUnitType
                                try
                                {
                                    DisplayUnitType dut = curParam.DisplayUnitType;
                                    if (dut != DisplayUnitType.DUT_UNDEFINED)
                                    {
                                        dValue = UnitUtils.ConvertToInternalUnits(dValue, dut);
                                    }
                                }
                                catch { }
#endif
                                curParam.Set(dValue);
                                return true;
                            }
                            return false;

                        case StorageType.Integer:
                            if (IsYesNoParameter(curParam))
                            {
                                // Handle Yes/No parameters
                                string lowerValue = value.ToLower();
                                if (lowerValue == "yes" || lowerValue == "true" || lowerValue == "1")
                                {
                                    curParam.Set(1);
                                }
                                else if (lowerValue == "no" || lowerValue == "false" || lowerValue == "0")
                                {
                                    curParam.Set(0);
                                }
                                else
                                {
                                    return false;
                                }
                                return true;
                            }
                            else
                            {
                                int iValue;
                                if (int.TryParse(value, out iValue))
                                {
                                    curParam.Set(iValue);
                                    return true;
                                }
                                return false;
                            }

                        case StorageType.String:
                            curParam.Set(value);
                            return true;

                        case StorageType.ElementId:
                            // For ElementId parameters, try to parse as integer
                            int idValue;
                            if (int.TryParse(value, out idValue))
                            {
                                curParam.Set(new ElementId(idValue));
                                return true;
                            }
                            // If not a number, try to find element by name
                            // This would require more context about what type of element to look for
                            return false;

                        default:
                            return false;
                    }
                }
                catch (Exception ex)
                {
                    Debug.Print($"Failed to set parameter value for {paramName}: {ex.Message}");
                    return false;
                }
            }
            return false;
        }

        // Helper method to check if a parameter is Yes/No type
        private static bool IsYesNoParameter(Parameter param)
        {
            // Check common Yes/No parameter names
            string paramName = param.Definition.Name.ToLower();
            if (paramName.Contains("yes") || paramName.Contains("no") ||
                paramName.Contains("structural") || paramName.Contains("bearing") ||
                paramName.Contains("enabled") || paramName.Contains("disabled"))
            {
                return true;
            }

            // Check if it's a known Yes/No built-in parameter
            if (param.Id.IntegerValue == (int)BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)
                return true;

            // For integer parameters with values only 0 or 1, assume Yes/No
            // This is a heuristic approach since we can't access ParameterType directly
            return false;
        }

        // This is a helper method to map a string name to a BuiltInParameter enum.
        internal static BuiltInParameter GetBuiltInParameterByName(string paramName)
        {
            // Common built-in parameters
            if (paramName == "Comments") return BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS;
            if (paramName == "Type Comments") return BuiltInParameter.ALL_MODEL_TYPE_COMMENTS;
            if (paramName == "Mark") return BuiltInParameter.ALL_MODEL_MARK;
            if (paramName == "Type Mark") return BuiltInParameter.ALL_MODEL_TYPE_MARK;
            if (paramName == "Family") return BuiltInParameter.ELEM_FAMILY_PARAM;
            if (paramName == "Family and Type") return BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM;
            if (paramName == "Type") return BuiltInParameter.ELEM_TYPE_PARAM;
            if (paramName == "Description") return BuiltInParameter.ALL_MODEL_DESCRIPTION;
            if (paramName == "Manufacturer") return BuiltInParameter.ALL_MODEL_MANUFACTURER;
            if (paramName == "Model") return BuiltInParameter.ALL_MODEL_MODEL;
            if (paramName == "URL") return BuiltInParameter.ALL_MODEL_URL;
            if (paramName == "Cost") return BuiltInParameter.ALL_MODEL_COST;
            if (paramName == "Assembly Code") return BuiltInParameter.UNIFORMAT_CODE;
            if (paramName == "Assembly Description") return BuiltInParameter.UNIFORMAT_DESCRIPTION;
            if (paramName == "Keynote") return BuiltInParameter.KEYNOTE_PARAM;

            // Wall parameters
            if (paramName == "Width") return BuiltInParameter.WALL_ATTR_WIDTH_PARAM;
            if (paramName == "Function") return BuiltInParameter.FUNCTION_PARAM;
            if (paramName == "Height") return BuiltInParameter.WALL_USER_HEIGHT_PARAM;
            if (paramName == "Base Offset") return BuiltInParameter.WALL_BASE_OFFSET;
            if (paramName == "Top Offset") return BuiltInParameter.WALL_TOP_OFFSET;

            // Floor parameters
            if (paramName == "Default Thickness") return BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM;
            if (paramName == "Thickness") return BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM;
            if (paramName == "Structural") return BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL;

            // Door/Window parameters
            if (paramName == "Head Height") return BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM;
            if (paramName == "Sill Height") return BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM;

            // Computed parameters
            if (paramName == "Area") return BuiltInParameter.HOST_AREA_COMPUTED;
            if (paramName == "Volume") return BuiltInParameter.HOST_VOLUME_COMPUTED;
            if (paramName == "Perimeter") return BuiltInParameter.HOST_PERIMETER_COMPUTED;
            if (paramName == "Level") return BuiltInParameter.LEVEL_PARAM;

            return BuiltInParameter.INVALID;
        }

        // This is a helper method to get the parameter storage type as a string.
        internal static string GetParameterStorageTypeString(StorageType storageType)
        {
            switch (storageType)
            {
                case StorageType.None:
                    return "None";
                case StorageType.Integer:
                    return "Integer";
                case StorageType.Double:
                    return "Double";
                case StorageType.String:
                    return "String";
                case StorageType.ElementId:
                    return "ElementId";
                default:
                    return "Unknown";
            }
        }
    }
}