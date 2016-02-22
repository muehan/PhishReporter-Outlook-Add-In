using System;

using Microsoft.Office.Tools.Ribbon;

namespace PhishingReporter
{
    using System.Windows.Forms;

    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Win32;

    using PhishingReporter.Properties;

    public partial class HOME
    {
        private void SecurityRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Phishing_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer exp = Globals.ThisAddIn.Application.ActiveExplorer();

            string germanText = "Diese markierte Nachricht wird weitergeleitet an "
                                + PhishReporterConfig.SecurityTeamEmailAlias + Environment.NewLine
                                + "und aus Ihrem Posteingang entfernt. Möchten Sie fortfahren?";

            string englishText = "The selected message will be forwarded to "
                                 + PhishReporterConfig.SecurityTeamEmailAlias + Environment.NewLine
                                 + " and removed from your inbox. Would you like to continue?";

            if (exp.Selection.Count == 1)
            {
                string messageBoxText = englishText;
                if (IsLanguageGerman())
                {
                    messageBoxText = germanText;
                }
                DialogResult response = MessageBox.Show(Resources.MessageBox_Title, messageBoxText, MessageBoxButtons.YesNo);
                if (response == DialogResult.Yes)
                {
                    // TODO: Be able to handle multiple selected messages rather than just the first one.
                    MailItem phishEmail = exp.Selection[0];
                    MailItem reportEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);

                    if ((IsFakeMail(phishEmail.HTMLBody)))
                    {
                        MessageBox.Show(Resources.MessageBox_Title, Resources.FakeMessage_Text, MessageBoxButtons.OK);
                        phishEmail.Delete();
                        return;
                    }

                    reportEmail.Attachments.Add(phishEmail, OlAttachmentType.olEmbeddeditem);
                    reportEmail.Subject = PhishReporterConfig.ReportEmailSubject + " - '" + phishEmail.Subject + "'" + new Random().Next(1000, 9999);
                    reportEmail.To = PhishReporterConfig.SecurityTeamEmailAlias;
                    reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the PhishReporter Outlook plugin. Please review the attached phishing email";

                    reportEmail.Send();
                    phishEmail.Delete();
                    MessageBox.Show(Resources.MessageBox_Title, Resources.DankeYouMsgBox_Text, MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show(Resources.MessageBox_Title, "Please Select a message To Continue.", MessageBoxButtons.OK);
            }
        }

        private bool IsFakeMail(string body)
        {
            Array fakeList = PhishReporterConfig.FakePhishingUrl.Split(',');

            foreach (string url in fakeList)
            {
                if (body.Contains(url))
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsLanguageGerman()
        {
            int languageCode = 1033; //Default to english 1031 for German
            int germanLanguageCode = 1031;

            const string keyEntry = "UILanguage";

            const string reg = @"Software\Microsoft\Office\14.0\Common\LanguageResources";
            try
            {
                RegistryKey k = Registry.CurrentUser.OpenSubKey(reg);
                if (k != null && k.GetValue(keyEntry) != null) languageCode = (int)k.GetValue(keyEntry);

            }
            catch { }

            try
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(reg);
                if (k != null && k.GetValue(keyEntry) != null) languageCode = (int)k.GetValue(keyEntry);
            }
            catch { }

            if (languageCode == germanLanguageCode)
            {
                return true;
            }

            return false;
        }
    }
}
