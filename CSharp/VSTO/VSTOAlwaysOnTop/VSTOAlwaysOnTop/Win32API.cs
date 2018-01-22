using xlApp = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace VSTOAlwaysOnTop
{
    public class Win32API
    {
        private const int SWP_NOMOVE = 0x0002;
        private const int SWP_NOSIZE = 0x0001;

        private const int HWND_BOTTOM = 1;
        private const int HWND_TOPMOST = -1;
        private const int HWND_NOTOPMOST = -2;

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        private IntPtr hwnd;
        private IntPtr xStype;

        public Win32API()
        {
            var getXLApp = GetExcelObject();
            hwnd = (IntPtr)getXLApp.Hwnd;
        }

        public void WindowOnTop(bool bln)
        {
            if(bln)
            {
                xStype = (IntPtr)HWND_TOPMOST;
            }
            else
            {
                xStype = (IntPtr)HWND_NOTOPMOST;
            }

            SetWindowPos(hwnd, xStype, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
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

        internal Process GetExcelProcess(xlApp.Application xlApp)
        {
            Process[] excelProcesses = Process.GetProcessesByName("Excel");
            foreach (Process excelProcess in excelProcesses)
            {
                if (xlApp.Hwnd == excelProcess.MainWindowHandle.ToInt32())
                {
                    return excelProcess;
                }
            }
            throw new InvalidOperationException(
               "Unexplained operation of the 'Process' class: the Excel process could not be returned.");
        }
    }
}
