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


namespace RAA_Level2
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class MyForm : Window
    {
        public MyForm()
        {
            InitializeComponent();
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

        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = @"C:\";
            //openFile.Filter = "csv files (*.csv) |*.csv";
            //openFile.Title = "Select a CSV file";
            openFile.Filter = "Excel files (*.xlsx;*.xls) |*.xlsx;*.xls";
            openFile.Title = "Select an Excel file";

            if (openFile.ShowDialog() == true)
            {
                txtFilePath.Text = openFile.FileName;
            }
            else 
            { 
                txtFilePath.Text = "";
            }
        }

        public string GetFilePath()
        {
            return txtFilePath.Text;
        }

        public string GetUnits()
        {
            if (rdoFeet.IsChecked == true)
            {
                return "feet";
            }
            else
            {
                return "mm";
            }
        }

        public bool GetPlansCreate()
        {
            if (chkFloorPlans.IsChecked == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool GetCeilingPlansCreate()
        {
            if (chkCeilingPlans.IsChecked == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
