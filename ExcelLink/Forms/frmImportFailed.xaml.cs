using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;
using System.IO;

namespace ExcelLink.Forms
{
    public partial class frmImportFailed : Window
    {
        private List<ImportErrorItem> _errorItems;

        public frmImportFailed(List<ImportErrorItem> errorItems)
        {
            InitializeComponent();
            _errorItems = errorItems;
            errorListView.ItemsSource = _errorItems;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void btnSaveReport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
            saveDialog.FileName = $"Report_{timestamp}.xlsx";
            saveDialog.Filter = "Excel Files|*.xlsx";

            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = saveDialog.FileName;

                Excel.Application excelApp = null;
                Excel.Workbook workbook = null;
                Excel.Worksheet worksheet = null;

                try
                {
                    excelApp = new Excel.Application();
                    workbook = excelApp.Workbooks.Add();
                    worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                    // Write headers
                    worksheet.Cells[1, 1] = "Element ID";
                    worksheet.Cells[1, 2] = "Description";

                    // Write data
                    for (int i = 0; i < _errorItems.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1] = _errorItems[i].ElementId;
                        worksheet.Cells[i + 2, 2] = _errorItems[i].Description;
                    }

                    worksheet.Columns.AutoFit();
                    workbook.SaveAs(filePath);
                    excelApp.Visible = true;

                    // Show frmInfoDialog instead of MessageBox
                    frmInfoDialog infoDialog = new frmInfoDialog("Report saved successfully.");
                    infoDialog.ShowDialog();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error saving report: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    if (workbook != null)
                    {
                        try { workbook.Close(false); } catch { }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    }
                    if (excelApp != null)
                    {
                        // Comment out or remove these two lines to keep Excel open
                        // try { excelApp.Quit(); } catch { }
                        // System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    }
                }
            }
        }
    }

    public class ImportErrorItem
    {
        public string ElementId { get; set; }
        public string Description { get; set; }
    }
}