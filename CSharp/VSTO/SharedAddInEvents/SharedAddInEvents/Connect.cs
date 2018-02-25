using Extensibility;
using Microsoft.Office.Core;
using stdole;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

// �� �߰� ����� ȣ��Ʈ ���� ���α׷��� COMAddIns �÷����� ���� ��ü�� �ڵ�ȭ�� �����մϴ�. 
// �� ��ü���� CreateCustomTaskPane�̶�� �� ���� �޼��尡 ������
// VBA�� �����Ͽ� ��� in-proc �Ǵ� �ܺ� �ڵ�ȭ Ŭ���̾�Ʈ���� ȣ�� �� �� �ֽ��ϴ�.

/*
' SharedAddInEvents.tlb
' mscorlib.tlb�� �����Ǿ�� addInUtils_SomeEvent ���� �߻� X

' Any CPU �����
' HKCR_Software_Microsoft_Office_Excel_Addins_SharedAddInEvents.Connect Ű �߰�

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
                // �� �߰� ��� ������� ���ǵ� ����� ���� UserControl�� ������� 
                // ����� ���� �۾� â�� ����ϴ�.
                CustomTaskPane taskPane =
                    factory.CreateCTP(
                    "SharedAddInEvents.SimpleUserControl",
                    "My Caption", Type.Missing);

                // Office���� ����� ���� �۾� â�� ����� ����� ���� UserControl�� 
                // �����ϰ� �߰� ��� ��ü �Ӽ��� ������ �� �ֽ��ϴ�.
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
        // ����� ���� �۾� â���� ���߸� Ŭ���ϸ� �̺�Ʈ�� UserControl���� 
        // �߰� ��� ��ü�� ���ĵǰ� ��� Ŭ���̾�Ʈ�� �̺�Ʈ�� ���ĵ˴ϴ�.
        // ���� �� COMAddIn.Object ��ü�� ���� �̺�Ʈ�� �����ϰ� �ֽ��ϴ�.
        internal void FireClickEvent(object sender, EventArgs e)
        {
            addInUtilities.FireEvent(sender, e);
        }
    }
}