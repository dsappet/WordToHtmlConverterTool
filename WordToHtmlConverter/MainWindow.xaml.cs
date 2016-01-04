using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordToHtmlConverter
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

        private void button_Click(object sender, RoutedEventArgs e)
        {
            //FileInfo file = new FileInfo(@"E:\Users\douglas\Documents\GitHub\WordToHtmlConverterTool\WordToHtmlConverter\Example.docx");
            FileInfo file = new FileInfo(FileToConvertTextBox.Text);

            if (file.Exists)
            {
                DirectoryInfo outputDir = new DirectoryInfo(Environment.CurrentDirectory + @"/outDir");
                if (!outputDir.Exists)
                {
                    Directory.CreateDirectory(outputDir.FullName);
                }

                var htmlFile = Converter.ToHtml(file.FullName, outputDir.FullName);

                var uri = new Uri(htmlFile.FullName);

                PreviewWindow.Navigate(uri);
            }
            else
            {
                MessageBox.Show("That file does not exist");
            }

        }

        private void FileToConvertTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".docx";
            dlg.Filter = "Word Document (*.docx) |*.docx";
            if (dlg.ShowDialog() == true)
            {
                FileToConvertTextBox.Text = dlg.FileName;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            FileToConvertTextBox.Text = @"E:\Users\douglas\Documents\GitHub\WordToHtmlConverterTool\WordToHtmlConverter\Example.docx";
        }
    }
}
