namespace PhishingReporter.CustomMessageBox
{
    using System.Windows.Forms;

    public class PhishingReporterMessageBox : Form
    {
        private Button ButtonPhishing;
        private Button ButtonSpam;
        private Button ButtonCancel;
        private Label lableText;

        public CustomDialogResult CustomDialogResult { get; set; }

        public PhishingReporterMessageBox()
        {
            InitializeComponent();
        }

        public DialogResult Show(string formTitle, string text)
        {
            lableText.Text = text;
            this.Text = formTitle;
            this.CenterToScreen();
            
            return this.ShowDialog();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PhishingReporterMessageBox));
            this.lableText = new System.Windows.Forms.Label();
            this.ButtonPhishing = new System.Windows.Forms.Button();
            this.ButtonSpam = new System.Windows.Forms.Button();
            this.ButtonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lableText
            // 
            this.lableText.AutoSize = true;
            this.lableText.Location = new System.Drawing.Point(13, 13);
            this.lableText.Name = "lableText";
            this.lableText.Size = new System.Drawing.Size(0, 13);
            this.lableText.TabIndex = 0;
            // 
            // ButtonPhishing
            // 
            this.ButtonPhishing.Location = new System.Drawing.Point(368, 118);
            this.ButtonPhishing.Name = "ButtonPhishing";
            this.ButtonPhishing.Size = new System.Drawing.Size(75, 23);
            this.ButtonPhishing.TabIndex = 1;
            this.ButtonPhishing.Text = "Phishing";
            this.ButtonPhishing.UseVisualStyleBackColor = true;
            this.ButtonPhishing.Click += new System.EventHandler(this.ButtonPhishing_Click);
            // 
            // ButtonSpam
            // 
            this.ButtonSpam.Location = new System.Drawing.Point(287, 118);
            this.ButtonSpam.Name = "ButtonSpam";
            this.ButtonSpam.Size = new System.Drawing.Size(75, 23);
            this.ButtonSpam.TabIndex = 2;
            this.ButtonSpam.Text = "Spam";
            this.ButtonSpam.UseVisualStyleBackColor = true;
            this.ButtonSpam.Click += new System.EventHandler(this.ButtonSpam_Click);
            // 
            // ButtonCancel
            // 
            this.ButtonCancel.Location = new System.Drawing.Point(12, 118);
            this.ButtonCancel.Name = "ButtonCancel";
            this.ButtonCancel.Size = new System.Drawing.Size(75, 23);
            this.ButtonCancel.TabIndex = 3;
            this.ButtonCancel.Text = "Cancel";
            this.ButtonCancel.UseVisualStyleBackColor = true;
            this.ButtonCancel.Click += new System.EventHandler(this.ButtonCancel_Click);
            // 
            // PhishingReporterMessageBox
            // 
            this.ClientSize = new System.Drawing.Size(455, 153);
            this.Controls.Add(this.ButtonCancel);
            this.Controls.Add(this.ButtonSpam);
            this.Controls.Add(this.ButtonPhishing);
            this.Controls.Add(this.lableText);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PhishingReporterMessageBox";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void ButtonCancel_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Dispose();
        }

        private void ButtonPhishing_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.OK;
            CustomDialogResult = CustomDialogResult.Phishing;
            Dispose();
        }

        private void ButtonSpam_Click(object sender, System.EventArgs e)
        {
            DialogResult = DialogResult.OK;
            CustomDialogResult = CustomDialogResult.Spam;
            Dispose();
        }
    }
}
