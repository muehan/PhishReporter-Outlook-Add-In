using System;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Office.Interop.Outlook;
using PhishingReporter.Properties;

namespace PhishingReporter
{
    public partial class SecurityRibbon
    {
        private void SecurityRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Phishing_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            string germanText = "Diese markierte Nachricht wird weitergeleitet an "
                                + PhishReporterConfig.SecurityTeamEmailAlias + Environment.NewLine
                                + "und aus Ihrem Posteingang entfernt. Möchten Sie fortfahren?";

            string englishText = "The selected message will be forwarded to "
                                 + PhishReporterConfig.SecurityTeamEmailAlias + Environment.NewLine
                                 + " and removed from your inbox. Would you like to continue?";

            string germanThankMessage = Resources.ThankYouMsgBox_Text_german;

            string englishTankMessage = Resources.ThankYouMsgBox_Text_english;

            string germanFakeMessageText = Resources.FakeMessage_Text_german;

            string englishFakeMessageText = Resources.FakeMessage_Text_english;

            string germanNoMessageText = Resources.NoMessageText_german;

            string englishNoMessageText = Resources.NoMessageText_english;

            string messageBoxText = englishText;
            string thanksMessageText = englishTankMessage;
            string fakeMessageText = englishFakeMessageText;
            string noMessageText = englishNoMessageText;

            if (IsLanguageGerman())
            {
                messageBoxText = germanText;
                thanksMessageText = germanThankMessage;
                fakeMessageText = germanFakeMessageText;
                noMessageText = germanNoMessageText;
            }

            if (explorer.Selection.Count == 1)
            {
                DialogResult response = MessageBox.Show(messageBoxText, Resources.MessageBox_Title, MessageBoxButtons.YesNo);
                if (response == DialogResult.Yes)
                {
                    // TODO: Be able to handle multiple selected messages rather than just the first one.
                    MailItem phishEmail = explorer.Selection[1];
                    MailItem reportEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);

                    if (IsFakeMail(phishEmail.HTMLBody))
                    {
                        MessageBox.Show(fakeMessageText, Resources.MessageBox_Title, MessageBoxButtons.OK);
                        phishEmail.Delete();
                        return;
                    }

                    reportEmail.Attachments.Add(phishEmail, OlAttachmentType.olEmbeddeditem);
                    reportEmail.Subject = PhishReporterConfig.ReportEmailSubject + " - '" + phishEmail.Subject + "'" + new Random().Next(1000, 9999);
                    reportEmail.To = PhishReporterConfig.SecurityTeamEmailAlias;
                    reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the PhishReporter Outlook plugin. Please review the attached phishing email";

                    reportEmail.Send();
                    phishEmail.Delete();

                    MessageBox.Show(thanksMessageText, Resources.MessageBox_Title, MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show(noMessageText, Resources.MessageBox_Title, MessageBoxButtons.OK);
            }
        }

        private bool IsFakeMail(string body)
        {
            Array fakeList = PhishReporterConfig.FakePhishingUrl.Split(',');

            foreach (string url in fakeList)
            {
                string trimmedUrl = url.Trim();

                if (body.Contains(trimmedUrl))
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
