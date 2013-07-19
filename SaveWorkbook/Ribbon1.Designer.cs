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
            this.btnSave = this.Factory.CreateRibbonButton();
            this.btnVMI = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnConfigure = this.Factory.CreateRibbonButton();
            this.btnSaveOAR = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnSave);
            this.group1.Items.Add(this.btnSaveOAR);
            this.group1.Items.Add(this.btnVMI);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnConfigure);
            this.group1.Label = "Save Workbook";
            this.group1.Name = "group1";
            // 
            // btnSave
            // 
            this.btnSave.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
            this.btnSave.Label = "Save Report";
            this.btnSave.Name = "btnSave";
            this.btnSave.ShowImage = true;
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
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
            // separator1
            // 
            this.separator1.Name = "separator1";
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
            // btnSaveOAR
            // 
            this.btnSaveOAR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveOAR.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveOAR.Image")));
            this.btnSaveOAR.Label = "Save Open AR";
            this.btnSaveOAR.Name = "btnSaveOAR";
            this.btnSaveOAR.ShowImage = true;
            this.btnSaveOAR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveOAR_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVMI;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConfigure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveOAR;
    }

    partial class ThisRibbonCollection
    {
        internal rbnSaveReport Ribbon1
        {
            get { return this.GetRibbon<rbnSaveReport>(); }
        }
    }
}
