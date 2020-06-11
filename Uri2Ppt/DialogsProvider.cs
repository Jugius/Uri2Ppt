using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Uri2Ppt
{
    static class DialogsProvider
    {
        public static string OpenExcelFile()
        {
            OpenFileDialog dlg = new OpenFileDialog { 
            Title = "Oткрыть файл Excel",
            Filter = "Файл Excel|*.xlsx"
            };
            if (dlg.ShowDialog() == true)
                return dlg.FileName;
            return null;
        }
        public static void ShowErrorMessage(string message, string caption)
            => ShowInfoMessage(message, caption, MessageBoxImage.Error);

        public static void ShowInfoMessage(string message, string caption, MessageBoxImage icon)
            => MessageBox.Show(message, caption, MessageBoxButton.OK, icon);

        public static void ShowException(Exception ex, string caption)
        {
            string message = ex.Message;
            if (ex.InnerException != null)
                message += "\nInner message:\n" + ex.InnerException.Message;
            if (!string.IsNullOrEmpty(ex.StackTrace))
                message += "\nStack:\n" + ex.StackTrace;
            if (!string.IsNullOrEmpty(ex.InnerException?.StackTrace))
                message += "\nInnerStack:\n" + ex.InnerException.StackTrace;
            ShowErrorMessage(message, caption ?? "Ошибка");
        }
    }
}
