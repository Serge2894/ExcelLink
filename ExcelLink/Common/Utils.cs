using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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

        internal static string GetParameterValueString(Element curElem, string paramName)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);
            if (curParam != null)
                return curParam.AsString();
            return string.Empty;
        }

        internal static string GetParameterValue(Element curElem, string paramName)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);
            if (curParam != null)
            {
                switch (curParam.StorageType)
                {
                    case StorageType.Double:
                        return curParam.AsDouble().ToString();
                    case StorageType.Integer:
                        return curParam.AsInteger().ToString();
                    case StorageType.String:
                        return curParam.AsString();
                    case StorageType.ElementId:
                        return curParam.AsElementId().IntegerValue.ToString();
                    default:
                        return string.Empty;
                }
            }
            return string.Empty;
        }

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
                            curParam.Set(double.Parse(value));
                            break;
                        case StorageType.Integer:
                            curParam.Set(int.Parse(value));
                            break;
                        case StorageType.String:
                            curParam.Set(value);
                            break;
                        case StorageType.ElementId:
                            // This is a complex case, not directly supported here.
                            // The value would need to be an ElementId.
                            return false;
                        default:
                            return false;
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    Debug.Print($"Failed to set parameter value for {paramName}: {ex.Message}");
                    return false;
                }
            }
            return false;
        }

        internal static BuiltInParameter GetBuiltInParameterByName(string paramName)
        {
            // This is a helper method to map a string name to a BuiltInParameter enum.
            // This is a simplified implementation and may not cover all cases.
            // In a more complete implementation, you might use a pre-populated dictionary
            // for performance.
            if (paramName == "Comments") return BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS;
            if (paramName == "Family") return BuiltInParameter.ELEM_FAMILY_PARAM;
            if (paramName == "Family and Type") return BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM;
            if (paramName == "Type") return BuiltInParameter.ELEM_TYPE_PARAM;
            if (paramName == "Type Name") return BuiltInParameter.SYMBOL_NAME_PARAM;
            // ... add other mappings as needed.
            return BuiltInParameter.INVALID;
        }

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