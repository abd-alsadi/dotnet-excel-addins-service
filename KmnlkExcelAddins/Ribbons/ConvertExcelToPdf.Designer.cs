namespace KmnlkExcelAddins.Ribbons
{
    partial class ConvertExcelToPdf : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ConvertExcelToPdf()
            : base(Globals.Factory.GetRibbonFactory())
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
            this.tab = this.Factory.CreateRibbonTab();
            this.group = this.Factory.CreateRibbonGroup();
            this.btn_SaveAsPdf = this.Factory.CreateRibbonButton();
            this.tab.SuspendLayout();
            this.group.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab
            // 
            this.tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab.Groups.Add(this.group);
            this.tab.Label = "TabAddIns";
            this.tab.Name = "tab";
            // 
            // group
            // 
            this.group.Items.Add(this.btn_SaveAsPdf);
            this.group.Label = "Kmnlk";
            this.group.Name = "group";
            // 
            // btn_SaveAsPdf
            // 
            this.btn_SaveAsPdf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_SaveAsPdf.Label = "Save As Pdf";
            this.btn_SaveAsPdf.Name = "btn_SaveAsPdf";
            this.btn_SaveAsPdf.ShowImage = true;
            this.btn_SaveAsPdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SaveAsPdf_Click);
            // 
            // ConvertExcelToPdf
            // 
            this.Name = "ConvertExcelToPdf";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ConvertExcelToPdf_Load);
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SaveAsPdf;
    }

    partial class ThisRibbonCollection
    {
        internal ConvertExcelToPdf ConvertExcelToPdf
        {
            get { return new ConvertExcelToPdf(); }
        }
    }
}
