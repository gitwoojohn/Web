using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using xlApp = Microsoft.Office.Interop.Excel;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaAlwaysOnTop
{
    public class BaseClass : IExcelAddIn
    {
        private const int SWP_NOMOVE = 0x0002;
        private const int SWP_NOSIZE = 0x0001;

        private const int HWND_BOTTOM = 1;
        private const int HWND_TOPMOST = -1;
        private const int HWND_NOTOPMOST = -2;

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr hWnd, String text, String caption, uint type);

        private static IntPtr hwnd;
        private static BaseClass win32API;
        private IntPtr xStype;

        [ComVisible(true)]
        public class Ribbon : ExcelRibbon
        {
            // ExcelDnaAlwaysOnTop-AddIn.dna 파일 없이 UI 사용 할때( 각 버전별 2007 - 2016 )
            //public override string GetCustomUI(string uiName)
            //{
            //    MessageBox(new IntPtr(0), "GetCustomUI called!", "MessageBox", 0);

            //    return
            // 엑셀 2007 버전용
            //            @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' >
            //    <ribbon>
            //      <tabs>
            //        <tab id='CustomTab' label='Always On Top'>
            //          <group id='SampleGroup' label='WindowTop On / Off'>
            //            <button id='Button1' label='On'  size='large' onAction='ButtonOn' tag='ShowHelloMessage' />
            //            <button id='Button2' label='Off'  size='large' onAction='ButtonOff'/>
            //          </group >
            //        </tab>
            //      </tabs>
            //    </ribbon>
            //  </customUI>";

            // 새로운 탭을 만들어서 리본 메뉴 추가( 이하 2010 이상 버전 )
            //    @"<customUI xmlns ='http://schemas.microsoft.com/office/2009/07/customui'>
            //        <ribbon>
            //            <tabs>
            //                <tab id='customTab' label='OnTop' insertAfterMso='TabView'>
            //                    <group id='AlwaysOnTop' label='Always On Top'>
            //                        <button id='button1' onAction='ButtonOn' imageMso='PictureBrightnessGallery' label='On' size='large' />
            //                        <separator id='separator1' />
            //                        <button id='button2' onAction='ButtonOff' imageMso='Risks' label='Off' size='large' />
            //                    </group >
            //                </tab >
            //            </tabs>
            //        </ribbon>
            //    </customUI>";

            // 홈 탭에 리본 메뉴 추가    
            //    @"<customUI xmlns ='http://schemas.microsoft.com/office/2009/07/customui'>
            //        <ribbon>
            //            <tabs>
            //                <tab idMso='TabHome'>
            //                    <group id='AlwaysOnTop' label='Always On Top'>
            //                        <button id='button1' onAction='ButtonOn' imageMso='PictureBrightnessGallery' label='On' size='large' />
            //                        <separator id='separator1' />
            //                        <button id='button2' onAction='ButtonOff' imageMso='Risks' label='Off' size='large' />
            //                    </group >
            //                </tab >
            //            </tabs>
            //        </ribbon>
            //    </customUI>";
            //}

            public void ButtonOn(Office.IRibbonControl control)
            {
                win32API.WindowOnTop(true);
            }
            public void ButtonOFF(Office.IRibbonControl control)
            {
                win32API.WindowOnTop(false);
            }
        }

        public void WindowOnTop(bool bln)
        {
            if (bln)
            {
                xStype = (IntPtr)HWND_TOPMOST;
            }
            else
            {
                xStype = (IntPtr)HWND_NOTOPMOST;
            }

            SetWindowPos( hwnd, xStype, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
        }

        internal xlApp.Application GetExcelObject()
        {
            try
            {
                return (xlApp.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                return null;
            }
        }

        public void AutoOpen()
        {
            var getXLApp = GetExcelObject();
            hwnd = (IntPtr)getXLApp.Hwnd;
            win32API = new BaseClass();
            //MessageBox(new IntPtr(0), "Hello Excel!", "Hello Dialog", 0);
        }

        public void AutoClose()
        {
            //MessageBox(new IntPtr(0), "Bye Excel!", "Hello Dialog", 0);
        }
    }
}
