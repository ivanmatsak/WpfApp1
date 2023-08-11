using Infragistics.Documents.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string path = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                path = openFileDialog.FileName;

            
            comboBox.Items.Add(path);

            
        }
        private void Button_Click_Save(object sender, RoutedEventArgs e)
        {
            spreadsheet.Workbook.Save(comboBox.SelectedValue.ToString());
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (var stream = File.Open(comboBox.SelectedValue.ToString(), FileMode.Open))
            {

                spreadsheet.Workbook = Workbook.Load(stream);

               
            }
        }
    }
}
