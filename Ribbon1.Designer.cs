namespace IM_IRS_Demo
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Valuate = this.Factory.CreateRibbonButton();
            this.find_PV = this.Factory.CreateRibbonButton();
            this.find_Portfolio_PV = this.Factory.CreateRibbonButton();
            this.find_VaR = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "IM Calc";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Valuate);
            this.group1.Items.Add(this.find_PV);
            this.group1.Items.Add(this.find_Portfolio_PV);
            this.group1.Items.Add(this.find_VaR);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // Valuate
            // 
            this.Valuate.Label = "Valuate";
            this.Valuate.Name = "Valuate";
            this.Valuate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Valuate_Click);
            // 
            // find_PV
            // 
            this.find_PV.Label = "Find PV";
            this.find_PV.Name = "find_PV";
            this.find_PV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.find_PV_Click);
            // 
            // find_Portfolio_PV
            // 
            this.find_Portfolio_PV.Label = "Find Portfolio PV";
            this.find_Portfolio_PV.Name = "find_Portfolio_PV";
            this.find_Portfolio_PV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.find_Portfolio_PV_Click);
            // 
            // find_VaR
            // 
            this.find_VaR.Label = "FVaR";
            this.find_VaR.Name = "find_VaR";
            this.find_VaR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.find_VaR_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Valuate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton find_PV;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton find_Portfolio_PV;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton find_VaR;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
