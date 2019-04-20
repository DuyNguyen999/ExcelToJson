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
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using Microsoft.Win32;
using Aspose.Cells;

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
            // load file to be converted
            var workbook = new Aspose.Cells.Workbook("D:\\C# Practice\\exltojson\\myExcel.xlsx");

            // save in different formats
            workbook.Save( "output.pdf", Aspose.Cells.SaveFormat.Pdf);

        }
    }
}
