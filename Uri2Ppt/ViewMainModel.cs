using System.ComponentModel;

namespace Uri2Ppt
{
    public class ViewMainModel : INotifyPropertyChanged
    {
        private string _columnText = "A";
        private string _columnHyperlink = "B";
        private string _columnURI1 = "C";
        private string _columnURI2 = "D";
        private string _columnURI3 = "E";
        private string _columnURI4 = "F";
        private string _openedFile;// = @"c:\Users\a.slusarev\Desktop\test.xlsx";
        private int _rowStart = 2;
        private int _rowFinish;
        public string ColumnText {
            get => _columnText;
            set {
                _columnText = value;
                OnPropertyChanged("ColumnText");
            }
        }
        public string ColumnHyperlink {
            get => _columnHyperlink;
            set {
                _columnHyperlink = value;
                OnPropertyChanged("ColumnHyperlink");
            }
        }
        
        public string ColumnURI1 {
            get => _columnURI1;
            set {
                _columnURI1 = value;
                OnPropertyChanged("ColumnURI1");
            }
        }
        public string ColumnURI2
        {
            get => _columnURI2;
            set
            {
                _columnURI2 = value;
                OnPropertyChanged("ColumnURI2");
            }
        }
        public string ColumnURI3
        {
            get => _columnURI3;
            set
            {
                _columnURI3 = value;
                OnPropertyChanged("ColumnURI3");
            }
        }
        public string ColumnURI4
        {
            get => _columnURI4;
            set
            {
                _columnURI4 = value;
                OnPropertyChanged("ColumnURI4");
            }
        }
        public string OpenedFile {
            get => _openedFile;
            set {
                _openedFile = value;
                OnPropertyChanged("OpenedFile");
                OnPropertyChanged("OpenedFileDescription");
            }
        }
        public string OpenedFileDescription =>
            string.IsNullOrEmpty(_openedFile) ? "Файл не открыт" : System.IO.Path.GetFileNameWithoutExtension(_openedFile);
        public int RowStart {
            get => _rowStart;
            set {
                _rowStart = value;
                OnPropertyChanged("RowStart");
            }
        }
        public int RowFinish
        {
            get => _rowFinish;
            set
            {
                _rowFinish = value;
                OnPropertyChanged("RowFinish");
            }
        }
        
        private void OnPropertyChanged(string v) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        public event PropertyChangedEventHandler PropertyChanged;
    }
}
