using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClearLines.Anakin.TaskPane.TreeView
{
    public class ExcelViewModel
    {
        private readonly Application excel;

        internal ExcelViewModel(Application excel)
        {
            Workbooks = new ObservableCollection<WorkbookViewModel>();

            this.excel = excel;
            ((AppEvents_Event)excel).NewWorkbook += AddWorkbook;
            excel.WorkbookOpen += AddWorkbook;
            excel.WorkbookActivate += ActiveWorkbookChanged;
            excel.SheetActivate += ActiveSheetActivate;

            var workbooks = excel.Workbooks;
            foreach (var workbook in workbooks)
            {
                if (workbook is Workbook book)
                {
                    var workbookViewModel = new WorkbookViewModel(book);
                    Workbooks.Add(workbookViewModel);
                }
            }
        }

        public ObservableCollection<WorkbookViewModel> Workbooks { get; }

        private void ActiveSheetActivate(object activatedSheet)
        {
            UpdateWorkbooks();
        }

        private void ActiveWorkbookChanged(Excel.Workbook activatedWorkbook)
        {
            UpdateWorkbooks();
        }

        public void UpdateWorkbooks()
        {
            var workbooks = excel.Workbooks;
            foreach (var workbookViewModel in Workbooks)
            {
                var workbookIsOpen = false;
                foreach (var workbook in workbooks)
                {
                    if (workbookViewModel.Workbook == workbook)
                    {
                        workbookIsOpen = true;
                        break;
                    }
                }

                if (workbookIsOpen == false)
                {
                    Workbooks.Remove(workbookViewModel);
                    break;
                }
                else
                {
                    workbookViewModel.UpdateDisplayProperties();
                    UpdateWorksheets(workbookViewModel);
                }
            }
        }

        private void UpdateWorksheets(WorkbookViewModel workbookViewModel)
        {
            var workbook = workbookViewModel.Workbook;
            var worksheets = workbook.Worksheets;
            try
            {
                foreach (var worksheetViewModel in workbookViewModel.Worksheets)
                {
                    var worksheetIsOpen = false;
                    foreach (var sheet in worksheets)
                    {
                        if (sheet is Worksheet worksheet)
                        {
                            if (worksheet == worksheetViewModel.Worksheet)
                            {
                                worksheetIsOpen = true;
                                break;
                            }
                        }
                    }

                    if (worksheetIsOpen == false)
                    {
                        workbookViewModel.Worksheets.Remove(worksheetViewModel);
                        worksheetViewModel.UpdateDisplayProperties();
                        break;
                    }
                    else
                    {
                        worksheetViewModel.UpdateDisplayProperties();
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

        private void AddWorkbook(Workbook newWorkbook)
        {
            var workbookViewModel = new WorkbookViewModel(newWorkbook);
            Workbooks.Add(workbookViewModel);
        }
    }
}

//private void RemoveClosedWorkbooksAndWorksheets(WorkbookViewModel workbookViewModel)
//{
//    var workbook = workbookViewModel.Workbook;
//    var worksheets = workbook.Worksheets;
//    foreach (var worksheetViewModel in workbookViewModel.Worksheets)
//    {
//        var worksheetIsOpen = false;
//        foreach (var sheet in worksheets)
//        {
//            if (sheet is Excel.Worksheet worksheet)
//            {
//                if (worksheet == worksheetViewModel.Worksheet)
//                {
//                    worksheetIsOpen = true;
//                    break;
//                }
//            }
//        }

//        if (worksheetIsOpen == false)
//        {
//            workbookViewModel.Worksheets.Remove(worksheetViewModel);
//        }
//        else
//        {
//            worksheetViewModel.UpdateDisplayProperties();
//        }
//    }
//}

//private void ActiveSheetChanged(object activatedSheet)
//{
//    RemoveClosedWorkbooksAndWorksheets();
//}

//private void ActiveWorkbookChanged(Excel.Workbook activatedWorkbook)
//{
//    RemoveClosedWorkbooksAndWorksheets();
//}

//private void PopulateWorkbooks()
//{
//    //var excel = Globals.ThisAddIn.Application;

//    //((Excel.AppEvents_Event)excel).NewWorkbook += AddWorkbook;
//    //excel.WorkbookOpen += AddWorkbook;
//    //excel.WorkbookActivate += ActiveWorkbookChanged;
//    //excel.SheetActivate += ActiveSheetChanged;

//    var workbooks = excel.Workbooks;
//    foreach (var workbook in workbooks)
//    {
//        if (workbook is Excel.Workbook book)
//        {
//            var workbookViewModel = new WorkbookViewModel(book);
//            Workbooks.Add(workbookViewModel);
//        }
//    }
//}


