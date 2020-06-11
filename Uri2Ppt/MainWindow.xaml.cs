using System;
using System.Windows;

namespace Uri2Ppt
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly ViewMainModel thisModel;
        public MainWindow()
        {
            InitializeComponent();
            ViewMainModel model = new ViewMainModel();
            this.DataContext = model;
            thisModel = model;
        }

        private async void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            string file = DialogsProvider.OpenExcelFile();
            if (!string.IsNullOrEmpty(file))
            {
                try
                {
                    int finishRow = await Downloader.GetLastRow(file);
                    thisModel.RowFinish = finishRow;
                }
                catch (Exception ex)
                {
                    DialogsProvider.ShowException(ex, "Ошибка чтения файла");
                }
            }
            else
                thisModel.RowFinish = 0;

            thisModel.OpenedFile = file;
        }

        private async void btnBegin_Click(object sender, RoutedEventArgs e)
        {
            string file = thisModel.OpenedFile;
            if (string.IsNullOrEmpty(file))
            {
                DialogsProvider.ShowErrorMessage("Файл не открыт", "Ошибка");
                return;
            }
            if (!System.IO.File.Exists(file))
            {
                DialogsProvider.ShowErrorMessage($"Файл не найден: {file}", "Ошибка");
                return;
            }

            UpdateProgress(0);
            var progress = new Progress<int>(UpdateProgress);
            try
            {
                btnBegin.IsEnabled = false;
                _ = await Downloader.ReadAndWrite(thisModel, progress);
            }
            catch (Exception ex)
            {
                DialogsProvider.ShowException(ex, "Error");
            }
            finally { 
                UpdateProgress(0);
                btnBegin.IsEnabled = true;
            }
            
        }
        private void UpdateProgress(int prog)
        {
            progressBar.Value = prog;
        }
    }
}
