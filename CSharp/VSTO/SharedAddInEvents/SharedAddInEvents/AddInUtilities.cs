using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SharedAddInEvents
{
    // COMAddIn.Object가 구현하는 사용자 지정 인터페이스입니다.
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("C2743EFC-AD90-47a6-B1DC-12E52C6E2FE7")]
    public interface IAddInUtilities
    {
        void CreateCustomTaskPane();
    }

    // 사용자 지정 이벤트의 대리자
    [ComVisible(false)]
    public delegate void SomeEventHandler(object sender, EventArgs e);     

    // Outgoing (source/event) interface.
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("B6B46A6E-76DB-4499-8C88-711108BF5C49")]
    public interface IAddInEvents
    {
        [DispId(1)]
        void SomeEvent(object sender, EventArgs e);
    }

    // COMAddIns 컬렉션을 통해 표시 할 객체입니다.
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F743C9A0-DDEF-49d5-AEAA-2E6798814C23")]
    [ComSourceInterfaces(typeof(IAddInEvents))]
    public class AddInUtilities : StandardOleMarshalObject, IAddInUtilities
    {        
        private Connect addInObject;

        // COM 클라이언트의 싱크를 연결
        public event SomeEventHandler SomeEvent;     

        internal AddInUtilities(Connect o)
        {
            addInObject = o;
        }

        // 이것은 호출하는 클라이언트에게 공개하는 인터페이스 메소드입니다.
        public void CreateCustomTaskPane()
        {
            addInObject.CreateCustomTaskPane();
        }

        // 추가 기능 클래스가 이벤트를 발생시키는 메서드를 제공합니다.
        internal void FireEvent(object sender, EventArgs e)
        {
            SomeEvent?.Invoke(sender, e);
        }
    }
}
