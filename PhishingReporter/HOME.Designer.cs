namespace PhishingReporter
{
    partial class HOME : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HOME()
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
            this.PhishingReporter = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Phishing = this.Factory.CreateRibbonButton();
            this.PhishingReporter.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // PhishingReporter
            // 
            this.PhishingReporter.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.PhishingReporter.ControlId.OfficeId = "TabMail";
            this.PhishingReporter.Groups.Add(this.group1);
            this.PhishingReporter.Label = "TabMail";
            this.PhishingReporter.Name = "PhishingReporter";
            this.PhishingReporter.Position = this.Factory.RibbonPosition.BeforeOfficeId("");
            // 
            // group1
            // 
            this.group1.Items.Add(this.Phishing);
            this.group1.Label = "ICT Security";
            this.group1.Name = "group1";
            this.group1.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps");
            // 
            // Phishing
            // 
            this.Phishing.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Phishing.Image = global::PhishingReporter.Properties.Resources.Logo___ICT_SEC;
            this.Phishing.Label = "Report Phishing";
            this.Phishing.Name = "Phishing";
            this.Phishing.OfficeImageId = "TrustCenter";
            this.Phishing.ScreenTip = "Report phishing emails";
            this.Phishing.ShowImage = true;
            this.Phishing.SuperTip = "Use this button to report suspicious emails to the Pilatus Information Security t" +
    "eam.";
            this.Phishing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Phishing_Click);
            // 
            // HOME
            // 
            this.Name = "HOME";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.PhishingReporter);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SecurityRibbon_Load);
            this.PhishingReporter.ResumeLayout(false);
            this.PhishingReporter.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonButton Phishing;
        protected internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab PhishingReporter;
    }

    partial class ThisRibbonCollection
    {
        internal HOME Home
        {
            get { return this.GetRibbon<HOME>(); }
        }
    }
}
