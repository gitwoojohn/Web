using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClearLines.Anakin.TaskPane.TreeView
{
    public class WorkbookViewModel : INotifyPropertyChanged
    {
        private string name;
        private string author;

        internal WorkbookViewModel(Excel.Workbook workbook)
        {
            Workbook = workbook;

            // 확장자 제외 파일명만 추출
            name = workbook.Name.Substring(0, workbook.Name.LastIndexOf("."));
            author = workbook.Author;

            workbook.NewSheet += AddSheet;
            Worksheets = new ObservableCollection<WorksheetViewModel>();

            var worksheets = workbook.Worksheets;
            foreach (var sheet in worksheets)
            {
                AddSheet(sheet);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public string Name
        {
            get => name;

            set
            {
                if (value != name)
                {
                    name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public string Author
        {
            get => author;

            set
            {
                if (value != author)
                {
                    author = value;
                    OnPropertyChanged(nameof(Author));
                }
            }
        }

        public string ImagePath => "TreeView/Workbook.bmp";

        public ObservableCollection<WorksheetViewModel> Worksheets { get; }

        internal Excel.Workbook Workbook { get; }

        internal void UpdateDisplayProperties()
        {
            try
            {
                Name = Workbook.Name.Substring(0, Workbook.Name.LastIndexOf("."));
                Author = Workbook.Author;
            }
            catch
            {
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void AddSheet(object newSheet)
        {
            if (newSheet is Excel.Worksheet worksheet)
            {
                var worksheetViewModel = new WorksheetViewModel(worksheet);
                Worksheets.Add(worksheetViewModel);
            }
        }
    }
}
