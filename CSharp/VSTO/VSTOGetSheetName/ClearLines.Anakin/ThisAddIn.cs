using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;
using ClearLines.Anakin.TaskPane;
using ClearLines.Anakin.TaskPane.TreeView;

namespace ClearLines.Anakin
{
    public partial class ThisAddIn
    {
        public CustomTaskPane TaskPane { get; private set; }
        public Dictionary<string, CustomTaskPane> WbCtp = new Dictionary<string, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WorkbookActivate += Application_WorkbookActivate;

            if (WbCtp.Count == 0)
            {
                var taskPaneView = new TaskPaneView();
                TaskPane = TaskPaneManager.GetTaskPane("A", "시트 목록", () => taskPaneView);
                TaskPane.Width = 250;
                TaskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
                TaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                TaskPane.DockPositionChanged += new EventHandler(TaskPane_DockPositionChanged);

                var excel = Application;
                var anakinViewModel = new AnakinViewModel(excel);
                var anakinView = taskPaneView.AnakinView;
                anakinView.DataContext = anakinViewModel;

                //WbCtp.Add( Wb.FullName, TaskPane);
            }
            else
            {
                //TaskPane = WbCtp;
            }
            //var taskPaneView = new TaskPaneView();
            ////TaskPane = TaskPaneManager.GetTaskPane("A", "시트 목록", () => new TaskPaneView());

            //TaskPane = TaskPaneManager.GetTaskPane("A", "시트 목록", () => taskPaneView);
            ////TaskPane = CustomTaskPanes.Add(taskPaneView, "시트 목록");

            ////TaskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
            //TaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            //TaskPane.DockPositionChanged += new EventHandler(TaskPane_DockPositionChanged);

            //var excel = Application;
            //var anakinViewModel = new AnakinViewModel(excel);
            //var anakinView = taskPaneView.AnakinView;
            //anakinView.DataContext = anakinViewModel;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            var wbCtp = WbCtp.Where(wb => wb.Key == Wb.FullName).FirstOrDefault().Value;
            if (wbCtp == null)
            {
                var taskPaneView = new TaskPaneView();
                TaskPane = TaskPaneManager.GetTaskPane("A", "시트 목록", () => taskPaneView);
                TaskPane.Width = 250;
                TaskPane.VisibleChanged += new EventHandler(TaskPane_VisibleChanged);
                TaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                TaskPane.DockPositionChanged += new EventHandler(TaskPane_DockPositionChanged);

                var excel = Application;
                var anakinViewModel = new AnakinViewModel(excel);
                var anakinView = taskPaneView.AnakinView;
                anakinView.DataContext = anakinViewModel;

                WbCtp.Add(Wb.FullName, TaskPane);
            }
            else
            {
                TaskPane = wbCtp;
            }
        }

        private void TaskPane_DockPositionChanged(object sender, EventArgs e)
        {
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.Ribbon.ShowHide.Checked = TaskPane.Visible;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
