using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace MergeExcel
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

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            foreach(String  s in   openFileDialog.FileNames)
            {
                    this.listBox.Items.Add(s);           
            }

        }

        private void btnDeleteFille_Click(object sender, RoutedEventArgs e)
        {
            this.listBox.Items.Remove(this.listBox.SelectedItem);
        }

        private void btnLoadFiles_Click(object sender, RoutedEventArgs e)
        {
            string[] files = new string[listBox.Items.Count];
            for (int i = 0; i < listBox.Items.Count; i++)
            {
                files[i] = listBox.Items[i].ToString();
            }

            var path = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);

            Excel.MergeExcel.DoMerge(files, path+@"\result.xlsx", "E", 2);
            DataTable oDataTable = Excel.MergeExcel.ViewData(path + @"\result.xlsx");
            this.dataGrid.ItemsSource = oDataTable.DefaultView;
        }

        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            var path = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            path = path.Substring(6);

            Excel.MergeExcel.openExcel( path + @"\result.xlsx");
        }
    }
}
