using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using XLApp = Microsoft.Office.Interop.Excel;

namespace AlwaysOnTop
{
    [ComVisible(true)]
    [Guid("A3BFF2AE-F8D7-42B0-B750-7581955A3CC4")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IWin32API
    {
        void WindowOnTop(bool bln, IntPtr hwnd);
    }

    [ComVisible(true)]
    [Guid("D19CB482-FEC4-4D8D-B9C2-504846985687")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Win32API : IWin32API
    {
        private const int SWP_NOMOVE = 0x0002;
        private const int SWP_NOSIZE = 0x0001;

        private const int HWND_BOTTOM = 1;
        private const int HWND_TOPMOST = -1;
        private const int HWND_NOTOPMOST = -2;

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        private IntPtr xStype;

        public void WindowOnTop(bool bln, IntPtr hwnd)
        {
            if (bln)
            {
                xStype = (IntPtr)HWND_TOPMOST;
            }
            else
            {
                xStype = (IntPtr)HWND_NOTOPMOST;
            }

            SetWindowPos(hwnd, xStype, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
        }
    }
}
