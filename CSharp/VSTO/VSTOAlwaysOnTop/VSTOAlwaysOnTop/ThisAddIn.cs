﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace VSTOAlwaysOnTop
{
    public partial class ThisAddIn
    {
        // 커스텀 XML 리본 메뉴 노출
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonX();
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

        public Excel.Worksheet GetActiveWorkSheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        #endregion
    }
}
