using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace DataPreparationToExcelWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static public string mUnits = "mm";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Select (object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog()
            {
                Multiselect = true,
                Title = "Select files",
            };
            dlg.Filter = "Output files|*.output";
            dlg.ShowDialog();
            if (dlg.FileName == String.Empty)
                return;
            string[] files_names = dlg.FileNames;
            //listBoxFiles.Items.Clear();
            //listBoxFiles.Items.Add("111");
            //listBoxFiles.ItemsSource = files_names;

            //ObservableCollection<string> oList;
            //oList = new System.Collections.ObjectModel.ObservableCollection<string>(files_names);
            //listBoxFiles.DataContext = oList;

            //Binding binding = new Binding();
            //listBox1.SetBinding(ListBox.ItemsSourceProperty, binding);
            foreach (string var in files_names) listBoxFiles.Items.Add(var);
            //listBoxFiles.Items.Add(files_names[0]);
            //listBoxFiles.Items.DeferRefresh();
            //listBoxFiles.Items.Add(files_names);
            //listBoxFiles.Items.AddRange(files_names);
            //ConverterToExcel.list.Clear();
            //ConverterToExcel.list.AddRange(files_names);
        }

        private void Button_Get(object sender, EventArgs e)
        {
            mUnits = cbUnits.Text;
            //ConverterToExcel.createExcelForListFiles(ConverterToExcel.list, mUnits);
            MessageBox.Show("Done");
        }
    }
}
