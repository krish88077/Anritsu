using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using ExcelLibrary.SpreadSheet;
using ExcelLibrary.CompoundDocumentFormat;
using System.IO;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Data;

namespace RTD
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Search");
            comboBox1.Items.Add("RTD Tool Details");
        }

        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                Search_tb.Visibility = System.Windows.Visibility.Visible;
                Search_btn.Visibility = System.Windows.Visibility.Visible;
                Search_lb.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                Search_lb.Visibility = System.Windows.Visibility.Visible;
                Search_tb.Visibility = System.Windows.Visibility.Hidden;
                Search_btn.Visibility = System.Windows.Visibility.Hidden;
            }
        }

        private void Search_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Search_tb.Text != "")
            {
                MessageBox.Show("Searching..........." + Search_tb.Text);
                Search_tb.Text = "";
            }
            else
            {
                MessageBox.Show("Enter some text");
            }
        }

        string path = @"../../RTD PRR.xlsx";
        string logtxt = "";

        private void ListBoxItem_Selected(object sender, RoutedEventArgs e)
        {
            dataGrid_disp.ItemsSource = null;
            
            string sheet = "Patch_Number_For_Component";
            string pathcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(pathcon);
            OleDbDataAdapter da = new OleDbDataAdapter("Select Component_name from [" + sheet + "$] ", conn);

            System.Data.DataTable dt = new System.Data.DataTable();

            da.Fill(dt);
            dataGrid_disp.ItemsSource = dt.DefaultView;
            dataGrid_disp.AutoGenerateColumns = true;

            logtxt = "Component name searching...";
            //Logs.Text = logtxt + DateTime.Now.ToString("hh:mm:ss tt");
            Logs.AppendText(Environment.NewLine + logtxt + DateTime.Now.ToString("hh:mm:ss tt"));
        }

        private void ListBoxItem_Selected_1(object sender, RoutedEventArgs e)
        {
            dataGrid_disp.ItemsSource = null;
            
            string sheet = "Patch_Number_For_Component";
            string pathcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(pathcon);
            OleDbDataAdapter da = new OleDbDataAdapter("Select Patch_Number_Allocation from [" + sheet + "$] ", conn);

            System.Data.DataTable dt1 = new System.Data.DataTable();

            da.Fill(dt1);
            dataGrid_disp.ItemsSource = dt1.DefaultView;
            dataGrid_disp.AutoGenerateColumns = true;

            logtxt = "RTD Patch series searching...";
            //Logs.Text = logtxt + DateTime.Now.ToString("hh:mm:ss tt");
            Logs.AppendText(Environment.NewLine + logtxt + DateTime.Now.ToString("hh:mm:ss tt"));
        }

        private void ListBoxItem_Selected_2(object sender, RoutedEventArgs e)
        {
            dataGrid_disp.ItemsSource = null;
           
            string sheet = "RTD_Patch_Releases _- _OPEN";
            string pathcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(pathcon);
            OleDbDataAdapter da = new OleDbDataAdapter("Select Patch_name from [" + sheet + "$] ", conn);

            System.Data.DataTable dt2 = new System.Data.DataTable();

            da.Fill(dt2);
            dataGrid_disp.ItemsSource = dt2.DefaultView;
            dataGrid_disp.AutoGenerateColumns = true;

            logtxt = "RTD Version searching...";
            //Logs.Text = logtxt + DateTime.Now.ToString("hh:mm:ss tt");
            Logs.AppendText(Environment.NewLine + logtxt + DateTime.Now.ToString("hh:mm:ss tt"));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            dataGrid_disp.ItemsSource = null;
           
            string sheet = "RTD_Patch_Releases_-_CLOSED";
            string pathcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(pathcon);
            OleDbDataAdapter da;

            if ((chkbx1.IsChecked == true) && (chkbx2.IsChecked == false) && (chkbx3.IsChecked == false) && (chkbx4.IsChecked == false))
            {
                da = new OleDbDataAdapter("Select Patch_name from [" + sheet + "$] ", conn);
            }
            else if ((chkbx1.IsChecked == true) && (chkbx2.IsChecked == true))
            {
                da = new OleDbDataAdapter("Select Patch_name , Customer from [" + sheet + "$] ", conn);
            }
            else if (chkbx1.IsChecked == true)
            {
                da = new OleDbDataAdapter("Select Patch_name from [" + sheet + "$] ", conn);
            }
            else if (chkbx2.IsChecked == true)
            {
                da = new OleDbDataAdapter("Select Customer from [" + sheet + "$] ", conn);
            }
            else if (chkbx3.IsChecked == true)
            {
                da = new OleDbDataAdapter("Select status from [" + sheet + "$] ", conn);
            }
            else if (chkbx4.IsChecked == true)
            {
                da = new OleDbDataAdapter("Select Created_by from [" + sheet + "$] ", conn);
            }
            else
                da = new OleDbDataAdapter("Select Created_by from [" + sheet + "$] ", conn);
            

            System.Data.DataTable dt3 = new System.Data.DataTable();

            da.Fill(dt3);
            dataGrid_disp.ItemsSource = dt3.DefaultView;
            dataGrid_disp.AutoGenerateColumns = true;

            logtxt = "Filtering the database...";
            //Logs.Text = logtxt + DateTime.Now.ToString("hh:mm:ss tt");
            Logs.AppendText(Environment.NewLine + logtxt + DateTime.Now.ToString("hh:mm:ss tt"));
        }

       
        
        private void Settings_icon_MouseDown(object sender, MouseButtonEventArgs e)
        {
            MessageBox.Show("Settings...");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataGrid_disp.ItemsSource = null;


            string sheet = "RTD_Patch_Releases_-_CLOSED";
            string pathcon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + path + ";Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(pathcon);
            OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + sheet + "$] ", conn);

            System.Data.DataTable dt4 = new System.Data.DataTable();

            da.Fill(dt4);
            dataGrid_disp.ItemsSource = dt4.DefaultView;
            dataGrid_disp.AutoGenerateColumns = true;

            if (dataGrid_disp.Columns.Count > 0)
                dataGrid_disp.Columns[dataGrid_disp.Columns.Count - 1].Width = new DataGridLength(1, DataGridLengthUnitType.Star); 
        }

       
    }
}
