using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClearLines.Anakin.TaskPane.TreeView
{
    public class WorksheetViewModel : INotifyPropertyChanged
    {
        private string name;
        public event PropertyChangedEventHandler PropertyChanged;

        internal Excel.Worksheet Worksheet { get; }

        public WorksheetViewModel(Excel.Worksheet worksheet)
        {
            Worksheet = worksheet;
            name = worksheet.Name;
        }

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
        public string ImagePath => "TreeView/Worksheet.bmp";

        internal void UpdateDisplayProperties()
        {
            try
            {
                Name = Worksheet.Name;
            }
            catch
            {
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
