using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SharedAddInEvents
{
    [ComVisible(true)]
    [Guid("6017B040-87CE-4bfc-88E8-18009A8EC403")]
    public partial class SimpleUserControl : UserControl
    {
        public Connect AddInObject;

        public SimpleUserControl()
        {
            InitializeComponent();
        }

        // When the user clicks the button on the taskpane,
        // we sink the Click event here, and fire the custom
        // event exposed from our IAddInUtilities object.
        private void btnHello_Click(object sender, EventArgs e)
        {
            AddInObject.FireClickEvent(sender, e);
        }
    }
}
