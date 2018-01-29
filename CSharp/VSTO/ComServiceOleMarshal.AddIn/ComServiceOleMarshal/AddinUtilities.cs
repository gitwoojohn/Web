using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace ComServiceOleMarshal
{
    [Serializable()]
    public class SomeObject
    {
        public int Number;
        public string Name;

        public SomeObject(int n, string s)
        {
            Number = n;
            Name = s;
        }
    }

    [ComVisible(true)]
    [Guid("F3140781-C4EA-45D6-92DE-A5D4FDF9ABB5")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAddInUtilities
    {
        void CurrentShowThread();
        void ImportData();
        void DisplayMessage();
        void SetCellValue(String cellAdress, object cellValue);
        void DoSomething(SomeObject o);
    }

    //[Serializable()]
    [ComVisible(true)]
    [Guid("19B832C5-D8CF-44D5-9987-20B115A8FE73")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : StandardOleMarshalObject, IAddInUtilities
    {
        public int threadID;

        // 테스트 메세지 표시
        public void DisplayMessage()
        {
            MessageBox.Show("Hello World");
        }

        // 스레드 확인
        public void CurrentShowThread()
        {
            int id = Thread.CurrentThread.ManagedThreadId;
            if (id == this.threadID)
            {
                MessageBox.Show("Same thread");
            }
            else
            {
                MessageBox.Show("Different thread");
            }

            //Globals.ThisAddIn.CreateNewTaskPane();
        }

        // 활성Sheet A1에 문자값 입력
        public void ImportData()
        {
            Excel.Worksheet wrkSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            wrkSheet.Range["A1"].Value = "This is ImportData";
            wrkSheet.Columns.AutoFit();
        }
        // 셀에 문자 날짜를 넣고 서식을 날짜 서식으로 변경
        public void SetCellValue(string cellAdress, object cellValue)
        {
            Excel.Worksheet wrkSheet = (Excel.Worksheet)
                Globals.ThisAddIn.Application.ActiveSheet;

            Excel.Range cell = wrkSheet.Range[cellAdress];
            cell.Value = cellValue;
            cell.NumberFormatLocal = "yyyy-mm-dd";
        }
        // TaskPane을 인수 입력 번호와 이름으로 변경해서 표시
        public void DoSomething(SomeObject o)
        {
            string captionTaskPane = $"{o.Number}={o.Name}";
            Globals.ThisAddIn.CreateNewTaskPane(captionTaskPane);
        }
    }
}