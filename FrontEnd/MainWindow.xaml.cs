using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MessageBox = System.Windows.MessageBox;

namespace FrontEnd
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        List<string> filePathsList = new List<string>();
        public MainWindow()
        {
            InitializeComponent();
        }


      
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (filePathsList.Count() <= 0)
            {
                const string caption = "Error";
                const string message = "No files selected to convert!";

                MessageBox.Show(message, caption);
            }

            else if (filePathsList.Count() > 5)
            {
                const string caption = "Error";
                const string message = "More than 5 files selected!";

                MessageBox.Show(message, caption);
            }

            else
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();

                string loc;
                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    loc = dialog.FileName;


                   
                    new Program(filePathsList, loc);
                 

                    MessageBox.Show("Process Complete", "Process Complete");

                    filePathsList.Clear();
                    lvDataBinding.ItemsSource = null;

                }

                
            }
           

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            openFileDlg.Multiselect = true;
            openFileDlg.Filter = "Office Files|*.xls;";


            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();
            // Get the selected file name and display in a TextBox.

            if (result == true)
            {
                

                foreach (string filePaths in openFileDlg.FileNames)
                {
                    filePathsList.Add(filePaths);
                }

               
                lvDataBinding.ItemsSource = filePathsList;
                
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            filePathsList.Clear();
            lvDataBinding.ItemsSource = null;
        }
    }
}
