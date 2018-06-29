using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClearLines.Anakin.TaskPane
{
    public partial class TaskPaneView : UserControl
    {
        public TaskPaneView()
        {
            InitializeComponent();
        }

        internal AnakinView AnakinView
        {
            get
            {
                return anakinView;
            }
        }

        private void TaskPaneView_SizeChanged(object sender, EventArgs e)
        {
            var taskPane = Globals.ThisAddIn.TaskPane;
            if (taskPane != null && taskPane.Width <= 250)
            {
                SendKeys.Send("{ESC}");
                Globals.ThisAddIn.TaskPane.Width = 250;
            }
        }
    }
}
