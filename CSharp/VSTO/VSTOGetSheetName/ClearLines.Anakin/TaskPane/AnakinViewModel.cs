using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClearLines.Anakin.TaskPane.TreeView;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClearLines.Anakin.TaskPane
{
    public class AnakinViewModel
    {
        private ExcelViewModel excelViewModel;
        private readonly Excel.Application excel;

        internal AnakinViewModel(Excel.Application excel)
        {
            this.excel = excel;
        }

        public ExcelViewModel ExcelViewModel
        {
            get
            {
                if (excelViewModel == null)
                {
                    excelViewModel = new ExcelViewModel(excel);
                }
                return excelViewModel;
            }
        }

        internal Excel.Workbook SelectedWorkbook
        {
            get;
            set;
        }

        internal Excel.Worksheet SelectedWorksheet
        {
            get;
            set;
        }
    }
}
