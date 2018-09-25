using ClearLines.Anakin.TaskPane.TreeView;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace ClearLines.Anakin.TaskPane
{
    /// <summary>
    /// AnakinView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class AnakinView : UserControl
    {
        public AnakinView()
        {
            InitializeComponent();
        }

        internal AnakinViewModel ViewModel
        {
            get
            {
                return DataContext as AnakinViewModel;
            }
        }

        private void treeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            //if (DataContext is AnakinViewModel modelWorkbook)
            //{
            //    if (e.NewValue is WorkbookViewModel workbookViewModel)
            //    {
            //        Debug.WriteLine($"NewValue : {e.NewValue.ToString()} //  OldValue : {e.OldValue.ToString()}");

            //        var workbook = workbookViewModel.Workbook;
            //        Debug.WriteLine(workbook.Name);
            //        workbook.Activate();
            //        modelWorkbook.SelectedWorkbook = workbook;
            //    }
            //    else
            //    {
            //        modelWorkbook.SelectedWorkbook = null;
            //    }
            //}

            if (DataContext is AnakinViewModel model)
            {
                if (e.NewValue is WorksheetViewModel worksheetViewModel)
                {
                    var worksheet = worksheetViewModel.Worksheet;
                    model.SelectedWorksheet = worksheet;
                    Debug.WriteLine(worksheet.Name);
                    worksheet.Activate();
                }
                else
                {
                    model.SelectedWorksheet = null;
                }
            }
        }
    }
}
