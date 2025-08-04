using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelLink.Forms
{
    /// <summary>
    /// Interaction logic for InfoDialog.xaml
    /// </summary>
    public partial class frmInfoDialog : Window
    {
        public frmInfoDialog()
        {
            InitializeComponent();
        }

        public frmInfoDialog(string message) : this()
        {
            txtInfoMessage.Text = message;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // This event handler will be called when the MouseDown event occurs on the Border
        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // If the left mouse button is pressed
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                // Drag the window
                DragMove();
            }
        }
    }
}