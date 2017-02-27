namespace PhishingReporter
{
    partial class SecurityRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SecurityRibbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonButton Phishing;
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            Phishing = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps");
            // 
            // group1
            // 
            this.group1.Items.Add(Phishing);
            this.group1.Label = "ICT Security";
            this.group1.Name = "group1";
            this.group1.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps");
            // 
            // Phishing
            // 
            Phishing.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            Phishing.Image = global::PhishingReporter.Properties.Resources.Logo___ICT_SEC;
            Phishing.Label = "Report Phishing";
            Phishing.Name = "Phishing";
            Phishing.OfficeImageId = "TrustCenter";
            Phishing.ScreenTip = "Report phishing emails";
            Phishing.ShowImage = true;
            Phishing.SuperTip = "Use this button to report suspicious emails to the Company Information Security team.";
            Phishing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Phishing_Click);
            // 
            // SecurityRibbon
            // 
            this.Name = "SecurityRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SecurityRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
    }

    partial class ThisRibbonCollection
    {
        internal SecurityRibbon SecurityRibbon
        {
            get { return this.GetRibbon<SecurityRibbon>(); }
        }
    }
}
