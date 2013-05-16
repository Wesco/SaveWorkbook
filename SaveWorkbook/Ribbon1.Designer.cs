namespace SaveWorkbook
{
    partial class rbnSaveReport : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rbnSaveReport() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(rbnSaveReport));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnGaps = this.Factory.CreateRibbonButton();
            this.btnVMI = this.Factory.CreateRibbonButton();
            this.btnISN117 = this.Factory.CreateRibbonButton();
            this.btn473 = this.Factory.CreateRibbonButton();
            this.btnConfigure = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnGaps);
            this.group1.Items.Add(this.btnVMI);
            this.group1.Items.Add(this.btnISN117);
            this.group1.Items.Add(this.btn473);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnConfigure);
            this.group1.Label = "Save Report";
            this.group1.Name = "group1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnGaps
            // 
            this.btnGaps.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGaps.Image = ((System.Drawing.Image)(resources.GetObject("btnGaps.Image")));
            this.btnGaps.Label = "GAPs";
            this.btnGaps.Name = "btnGaps";
            this.btnGaps.ShowImage = true;
            this.btnGaps.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGaps_Click);
            // 
            // btnVMI
            // 
            this.btnVMI.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnVMI.Image = ((System.Drawing.Image)(resources.GetObject("btnVMI.Image")));
            this.btnVMI.Label = "VMI";
            this.btnVMI.Name = "btnVMI";
            this.btnVMI.ShowImage = true;
            this.btnVMI.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVMI_Click);
            // 
            // btnISN117
            // 
            this.btnISN117.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnISN117.Image = ((System.Drawing.Image)(resources.GetObject("btnISN117.Image")));
            this.btnISN117.Label = "117";
            this.btnISN117.Name = "btnISN117";
            this.btnISN117.ShowImage = true;
            this.btnISN117.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnISN117_Click);
            // 
            // btn473
            // 
            this.btn473.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn473.Image = ((System.Drawing.Image)(resources.GetObject("btn473.Image")));
            this.btn473.Label = "473";
            this.btn473.Name = "btn473";
            this.btn473.ShowImage = true;
            this.btn473.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn473_Click);
            // 
            // btnConfigure
            // 
            this.btnConfigure.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConfigure.Image = ((System.Drawing.Image)(resources.GetObject("btnConfigure.Image")));
            this.btnConfigure.Label = "Configure";
            this.btnConfigure.Name = "btnConfigure";
            this.btnConfigure.ShowImage = true;
            this.btnConfigure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConfigure_Click);
            // 
            // rbnSaveReport
            // 
            this.Name = "rbnSaveReport";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGaps;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnISN117;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn473;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVMI;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConfigure;
    }

    partial class ThisRibbonCollection
    {
        internal rbnSaveReport Ribbon1
        {
            get { return this.GetRibbon<rbnSaveReport>(); }
        }
    }
}
