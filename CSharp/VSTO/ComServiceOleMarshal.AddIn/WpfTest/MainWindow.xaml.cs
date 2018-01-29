//#define DEBUG
//#define TRACE

using System;
using System.Windows;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;

namespace WpfTest
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        Excel.Application XLApp = null;
        ComServiceOleMarshal.IAddInUtilities utils = null;
        public MainWindow()
        {
            InitializeComponent();

            // 필터링 ( reject callee )
            // This registers the IOleMessageFilter to handle any threading errors.
            //MessageFilter.Register();

        }
        //private Excel.Application XLApp;

        private void ButtonInvokeAddIn_Click(object sender, RoutedEventArgs e)
        {
            // Creates the text file that the trace listener will write to.
            //FileStream myTraceLog = new FileStream( @"C:\myTraceLog.txt", FileMode.OpenOrCreate );
            // Creates the new trace listener.
            //TextWriterTraceListener myListener = new TextWriterTraceListener( myTraceLog );

            try
            {
                XLApp = new Excel.Application()
                {
                    Visible = true
                };

                //Thread.Sleep( 3000 );
                XLApp.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
                Object addInName = "ComServiceOleMarshal";
                Office.COMAddIn addIn = XLApp.COMAddIns.Item(ref addInName);

                while (utils == null)
                {
                    Thread.Sleep(100);
                    utils = (ComServiceOleMarshal.IAddInUtilities)addIn.Object;
                }

                utils.CurrentShowThread();

                ComServiceOleMarshal.SomeObject o = new ComServiceOleMarshal.SomeObject(123, "Hello");
                utils.DoSomething(o);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                // Flushes the buffers of all listeners in the Listeners collection.
                //Trace.Flush();
                //Flushes only the buffer of myListener.
                //myListener.Flush();
                //myTraceLog.Close();
                //And at the end ( 필터링 : reject callee )
                //MessageFilter.Revoke();

                //if (XLApp != null)
                //{
                //    XLApp.Quit();
                //    XLApp = null;
                //    utils = null;
                //}
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
            }
        }
        private void Un_Unloaded(object sender, RoutedEventArgs e)
        {
            if (XLApp != null)
            {
                XLApp.Quit();
                XLApp = null;
                utils = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        // Type 알아내기
        public static Type GetTypeFromComObject(object ComObject, Assembly OwnerAssembly)
        {

            // COM Object가 아니라면 원래 Type을 리턴합니다.
            if (ComObject.GetType().FullName != "System.__ComObject") return ComObject.GetType();

            // COM Object의 IUnknown 객체를 가져옵니다.
            // 모든 COM Object는 IUnknown을 상속받아 구현되며, 
            // 이는 추후 Com Object에 QueryInterface를 호출되기 위해 필요합니다.
            IntPtr IUnknown = Marshal.GetIUnknownForObject(ComObject);

            // 해당 Assembly에 정의된 모든 Type을 가져옵니다.
            Type[] Types = OwnerAssembly.GetTypes();

            foreach (Type Type in Types)
            {
                Guid TypeGUID = Type.GUID;
                if (Type.IsInterface == false || TypeGUID == Guid.Empty)
                {
                    // COM Interop Type은 반드시 Interface형식으로 사용되며,
                    // 각자 고유의 GUID를 갖기때무에 이에 해당하지 않은 Type은 스킵.
                    continue;
                }

                IntPtr InstancePointer = IntPtr.Zero;
                if (Marshal.QueryInterface(IUnknown, ref TypeGUID, out InstancePointer) == 0)
                {
                    // QueryInterface에 현재 확인중인 Type의 GUID를 이용하여,
                    // IUnknown 객체가 해당 Type의 명령을 수행 할 수 있는지 여부를 확인합니다.
                    return Type;
                }
            }
            return null;
        }
    }
}
//Microsoft KB문서( http://support.microsoft.com/kb/320523)에 보면 as 연산자를 이용해 
//하나씩 캐스팅을 할수 있다고 하지만, 이렇게 할경우 Type이 많아질 수록 
//더 많은 체크 루틴을 작성해야 하기 때문에 적합한 방법은 아닌것 같습니다.

//이번시간에 소개할 내용은 이러한 상황에서 COM객체의 Type을 확인 할 수 있는 방법에대해 소개합니다.
//GetTypeFromComObject Method
//COM 객체의 Type을 가져오기 위해서 몇 가지 COM객체의 특성을 이용합니다. 
//첫째로 COM객체는 GetType메서드를 호출하면 System.__ComObject타입을 리턴한다는 점과, 
//둘째로 COM Interop Type은 Interface이고, 고유의 GUID를 갖는다는 점을 이용합니다.
//소스코드에 대한 설명은 주석으로 대체합니다.