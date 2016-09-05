namespace PhishingReporter
{
    using System;
    using Microsoft.Office.Tools.Ribbon;
    using System.Windows.Forms;
    using Microsoft.Office.Interop.Outlook;

    using PhishingReporter.CustomMessageBox;
    using PhishingReporter.Properties;
    using PhishingReporter.Helpers;

    public partial class SecurityRibbon
    {
        private Language lang;

        private void SecurityRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            lang = LanguageHelper.GetLanguage();
        }

        private void Phishing_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer.Selection.Count == 1)
            {
                var dialog = new PhishingReporterMessageBox();
                dialog.Show(TextHelper.MessageBoxText(lang));
                if (dialog.DialogResult == DialogResult.OK && dialog.CustomDialogResult == CustomDialogResult.Phishing)
                {
                    // TODO: Be able to handle multiple selected messages rather than just the first one.
                    MailItem phishEmail = explorer.Selection[1];
                    MailItem reportEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);

                    if (IsFakeMail(phishEmail.HTMLBody))
                    {
                        MessageBox.Show(TextHelper.FakeMessageText(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);
                        phishEmail.Delete();
                        return;
                    }

                    reportEmail.Attachments.Add(phishEmail, OlAttachmentType.olEmbeddeditem);
                    reportEmail.Subject = PhishReporterConfig.ReportEmailSubject + " - '" + phishEmail.Subject + "'" + new Random().Next(1000, 9999);
                    reportEmail.To = PhishReporterConfig.SecurityTeamEmailAlias;
                    reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the PhishReporter Outlook plugin. Please review the attached phishing email";

                    reportEmail.Send();
                    phishEmail.Delete();

                    MessageBox.Show(TextHelper.ThankYouMessage(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);
                }
                else if (dialog.DialogResult == DialogResult.OK && dialog.CustomDialogResult == CustomDialogResult.Spam)
                {
                    MailItem phishEmail = explorer.Selection[1];
                    phishEmail.Move(Globals.ThisAddIn.GetSpamFolder());
                }
            }
            else
            {
                MessageBox.Show(TextHelper.NoEmailSelectedText(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);
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
    }
}
