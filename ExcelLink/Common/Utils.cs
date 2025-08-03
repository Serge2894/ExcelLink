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
                            curParam.Set(double.Parse(value));
                            break;
                        case StorageType.Integer:
                            curParam.Set(int.Parse(value));
                            break;
                        case StorageType.String:
                            curParam.Set(value);
                            break;
                        case StorageType.ElementId:
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

        // This is a helper method to map a string name to a BuiltInParameter enum.
        internal static BuiltInParameter GetBuiltInParameterByName(string paramName)
        {
            if (paramName == "Comments") return BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS;
            if (paramName == "Family") return BuiltInParameter.ELEM_FAMILY_PARAM;
            if (paramName == "Family and Type") return BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM;
            if (paramName == "Type") return BuiltInParameter.ELEM_TYPE_PARAM;
            if (paramName == "Type Name") return BuiltInParameter.SYMBOL_NAME_PARAM;

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