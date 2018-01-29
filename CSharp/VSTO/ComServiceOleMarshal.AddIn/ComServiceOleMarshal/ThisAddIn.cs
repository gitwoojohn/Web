using System;
using System.Threading;
using System.Windows.Forms;

/*
추가 기능에서 호스트의 COMAddIns 컬렉션을 통해 개체를 노출 할 수있게 해주는 
RequestComAddInAutomationService를 재정의합니다. 

이렇게 하면 COM개체에 VBA를 비롯한 다른 클라이언트에서 사용할 수 있습니다. 
*/

/*
' 
Sub VSTO_Call_VBA()
    Dim addin As Office.COMAddIn
    Dim automationObject As Object

    Set addin = Application.COMAddIns("ComServiceOleMarshal")
    Set automationObject = addin.Object

    automationObject.DisplayMessage
    automationObject.ImportData
End Sub
*/

namespace ComServiceOleMarshal
{
    public partial class ThisAddIn
    {
        private AddInUtilities utilities;

        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
            {
                utilities = new AddInUtilities()
                {
                    threadID = Thread.CurrentThread.ManagedThreadId
                };
            }
            return utilities;
        }
        internal void CreateNewTaskPane(string caption)
        {
            Microsoft.Office.Tools.CustomTaskPane taskPane =
                this.CustomTaskPanes.Add(new UserControl(), caption);
            taskPane.Visible = true;
        }
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
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
