using Extensibility;
using Microsoft.Office.Core;
using stdole;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

// 이 추가 기능은 호스트 응용 프로그램의 COMAddIns 컬렉션을 통해 개체를 자동화에 노출합니다. 
// 이 개체에는 CreateCustomTaskPane이라는 한 가지 메서드가 있으며
// VBA를 포함하여 모든 in-proc 또는 외부 자동화 클라이언트에서 호출 할 수 있습니다.

/*
' SharedAddInEvents.tlb
' mscorlib.tlb가 참조되어야 addInUtils_SomeEvent 에러 발생 X

' Any CPU 빌드시
' HKCR_Software_Microsoft_Office_Excel_Addins_SharedAddInEvents.Connect 키 추가

Public WithEvents addInUtils As SharedAddInEvents.AddinUtilities

Private Sub CommandButton1_Click()
    Dim addin As Office.COMAddIn

    Set addin = Application.COMAddIns("SharedAddInEvents.Connect")
    Set addInUtils = addin.Object
    addInUtils.CreateCustomTaskPane
End Sub

Private Sub addInUtils_SomeEvent( ByVal sender As Variant, ByVal e As mscorlib.EventArgs)
    MsgBox "Got SomeEvent"
End Sub
*/

namespace SharedAddInEvents
{
    [Guid("CA2575D4-E119-4257-B180-2121720CF773")]
    [ProgId("SharedAddInEvents.Connect")]
    public class Connect : object, IDTExtensibility2, ICustomTaskPaneConsumer
    {
        #region Fields

        private ICTPFactory factory;
        private AddInUtilities addInUtilities;

        private Excel.Application applicationObject = null;
        private object addInInstance;

        #endregion


        #region Standard IDTExtensibility2 methods

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                IPictureDisp ipd = applicationObject.CommandBars.GetImageMso("Paste", 32, 32);

                IconViewForm form = new IconViewForm();

                form.IconPictureBox.Image = IPictureDispConverter.ConvertPixelByPixel(ipd);

                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Unhandled Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion


        public void OnConnection(
            object application, 
            ext_ConnectMode connectMode,  
            object addInInst, ref Array custom)
		{
            addInUtilities = new AddInUtilities(this);
            COMAddIn comAddIn = (COMAddIn)addInInst;
            comAddIn.Object = addInUtilities;

            applicationObject = (Excel.Application)application;
            addInInstance = addInInst;

            if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
            {
                OnStartupComplete(ref custom);
            }
        }

        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            factory = CTPFactoryInst;
        }

        internal void CreateCustomTaskPane()
        {
            try
            {
                // 이 추가 기능 어셈블리에 정의된 사용자 지정 UserControl을 기반으로 
                // 사용자 지정 작업 창을 만듭니다.
                CustomTaskPane taskPane =
                    factory.CreateCTP(
                    "SharedAddInEvents.SimpleUserControl",
                    "My Caption", Type.Missing);

                // Office에서 사용자 지정 작업 창을 만들면 사용자 지정 UserControl을 
                // 추출하고 추가 기능 개체 속성을 설정할 수 있습니다.
                SimpleUserControl sc = 
                    (SimpleUserControl)taskPane.ContentControl;
                sc.AddInObject = this;
                taskPane.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        // 사용자 지정 작업 창에서 단추를 클릭하면 이벤트가 UserControl에서 
        // 추가 기능 개체로 전파되고 모든 클라이언트에 이벤트가 전파됩니다.
        // 노출 된 COMAddIn.Object 개체에 대해 이벤트를 수신하고 있습니다.
        internal void FireClickEvent(object sender, EventArgs e)
        {
            addInUtilities.FireEvent(sender, e);
        }
    }
}