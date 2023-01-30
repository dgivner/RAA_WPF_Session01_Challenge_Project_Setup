using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace RAA_WPF_Session01_Challenge_Project_Setup
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class MyWindow : Window
    {
        public MyWindow()
        {
            InitializeComponent();
        }
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = @"C:\";
            openFile.Filter = "csv files (*.csv)|*.csv";

            if (openFile.ShowDialog() == true)
            {
                tbxFile.Text = openFile.FileName;
            }
            else
            {
                tbxFile.Text = "";
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        public string GetTextBoxValue()
        {
            return tbxFile.Text;
        }
        public bool GetCheckbox1()
        {
            if (chbCheck1.IsChecked == true)
                return true;
            else
                return false;
        }

        public string GetGroup1()
        {
            if (rb1.IsChecked == true)
                return rb1.Content.ToString();
            else if (rb2.IsChecked == true)
                return rb2.Content.ToString();
            return null;
        }

    }
}
