using Largesky.Excel;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace Com.Largesky.App
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
                if (ofd.ShowDialog().Value == false)
                {
                    return;
                }
                var read = XlsxFileReader.Open(ofd.FileName);
                var data = read.ReadAllRows().ToList();
                DataTable dt = new DataTable();
                dt.Columns.AddRange(data[0].Select(obj => new DataColumn { ColumnName = obj }).ToArray());
                dt.Columns[0].ColumnName = "行号";
                for (int i = 1; i < data.Count; i++)
                {
                    dt.Rows.Add(data[i]);
                }
                this.dgvData.ItemsSource = dt.AsDataView();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnTest2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string file = @"D:\test.xlsx";
                string[][] contents = new string[1][];
                contents[0] = new string[3];
                contents[0][0] = "fdsfds";
                contents[0][1] = "123456";
                contents[0][2] = "你好";
                XlsxFileWriter.WriteXlsx(file, contents);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
