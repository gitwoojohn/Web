namespace ClearLines.Anakin
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.TaskPaneSample = this.Factory.CreateRibbonTab();
            this.toggleGroup = this.Factory.CreateRibbonGroup();
            this.ShowHide = this.Factory.CreateRibbonToggleButton();
            this.TaskPaneSample.SuspendLayout();
            this.toggleGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // TaskPaneSample
            // 
            this.TaskPaneSample.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TaskPaneSample.ControlId.OfficeId = "TabDeveloper";
            this.TaskPaneSample.Groups.Add(this.toggleGroup);
            this.TaskPaneSample.Label = "TabDeveloper";
            this.TaskPaneSample.Name = "TaskPaneSample";
            // 
            // toggleGroup
            // 
            this.toggleGroup.Items.Add(this.ShowHide);
            this.toggleGroup.Label = "작업창 토글";
            this.toggleGroup.Name = "toggleGroup";
            // 
            // ShowHide
            // 
            this.ShowHide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ShowHide.Image = ((System.Drawing.Image)(resources.GetObject("ShowHide.Image")));
            this.ShowHide.Label = "작업창 On-Off";
            this.ShowHide.Name = "ShowHide";
            this.ShowHide.ShowImage = true;
            this.ShowHide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowHide_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TaskPaneSample);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.TaskPaneSample.ResumeLayout(false);
            this.TaskPaneSample.PerformLayout();
            this.toggleGroup.ResumeLayout(false);
            this.toggleGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TaskPaneSample;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup toggleGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ShowHide;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
