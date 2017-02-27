using System.Diagnostics.PerformanceData;
using System.Runtime.InteropServices;
using System.Threading;

namespace PhishingReporter
{
    using System;
    using Microsoft.Office.Tools.Ribbon;
    using System.Windows.Forms;
    using Microsoft.Office.Interop.Outlook;
    using Properties;
    using CustomMessageBox;
    using Helpers;

    public partial class SecurityRibbon
    {
        private Language lang;

        private void SecurityRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            lang = Globals.ThisAddIn.GetUserLanguage();
        }

        /// <summary>
        /// OnClick here!
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Phishing_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer.Selection.Count != 1)
            {
                // Nothing selected
                return;
            }

            // Get selected email
            MailItem phishEmail = explorer.Selection[1];

            if (HandleFakeMail(phishEmail))
            {
                return;
            }

            if (HandleInternalMails(phishEmail.SenderEmailAddress))
            {
                return;
            }

            var emailType = PreCheckNewsletterOrSpam(phishEmail);

            if (emailType == EmailType.None)
            {
                emailType = AskUserIfSpamOrPhising(phishEmail);
            }

            HandleMail(emailType, phishEmail);
        }

        private void HandleMail(EmailType emailType, MailItem phishEmail)
        {
            switch (emailType)
            {
                case EmailType.Spam:

                    var spamDialogReslut = MessageBox.Show(TextHelper.SpamAddedText(lang), Resources.MessageBox_Title, MessageBoxButtons.OKCancel);

                    if (spamDialogReslut == DialogResult.OK)
                    {
                        phishEmail.Move(Globals.ThisAddIn.GetSpamFolder());
                        Globals.ThisAddIn.CreateJunkRule(phishEmail.SenderEmailAddress);
                    }
                    break;

                case EmailType.Phishing:
                    MailItem reportEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    string subject = PhishReporterConfig.ReportEmailSubject + " - " + phishEmail.Subject + " - " + new Random().Next(1000, 9999);

                    reportEmail.Attachments.Add(phishEmail, OlAttachmentType.olEmbeddeditem);
                    reportEmail.Subject = subject;
                    reportEmail.To = PhishReporterConfig.SecurityTeamEmailAlias;
                    reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the PhishReporter Outlook plugin. Please review the attached phishing email";
                    reportEmail.Send();

                    // Delete Mail Permanently
                    string phishSubject = phishEmail.Subject;
                    phishEmail.Delete();
                    DeleteEmailFromFolder(phishSubject, OlDefaultFolders.olFolderDeletedItems);

                    MessageBox.Show(TextHelper.ThankYouMessage(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);

                    // Delete Sent Mail
                    DeleteEmailFromFolder(subject, OlDefaultFolders.olFolderSentMail);
                    DeleteEmailFromFolder(subject, OlDefaultFolders.olFolderDeletedItems);
                    break;

                case EmailType.None:
                    //MessageBox.Show(TextHelper.NoEmailSelectedText(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);
                    break;
            }
        }

        private EmailType AskUserIfSpamOrPhising(MailItem phishEmail)
        {
            var dialog = new PhishingReporterMessageBox();
            dialog.Show("Phishing Reporter", TextHelper.MessageBoxText(lang));
            if (dialog.DialogResult == DialogResult.OK && dialog.CustomDialogResult == CustomDialogResult.Phishing)
            {
                return EmailType.Phishing;
                // TODO: Be able to handle multiple selected messages rather than just the first one.
            }

            if (dialog.DialogResult == DialogResult.OK && dialog.CustomDialogResult == CustomDialogResult.Spam)
            {
                return EmailType.Spam;
            }

            return EmailType.None;
        }

        private EmailType PreCheckNewsletterOrSpam(MailItem phishEmail)
        {
            string subject = phishEmail.Subject;

            subject = subject.ToLower().Trim();

            if (subject.Contains("spam") || subject.Contains("newsletter"))
            {
                return EmailType.Spam;
            }

            if (subject.Contains("phishing"))
            {
                return EmailType.Phishing;
            }

            return EmailType.None;
        }

        /// <summary>
        /// Check if email is comming from other Company Employee
        /// Employees can not be marked ad SPAM or Phinging
        /// </summary>
        /// <param name="emailAddress"></param>
        private bool HandleInternalMails(string emailAddress)
        {
            // Company internal mails contains no sender address, only exchange OU path
            if (!emailAddress.Contains("@"))
            {
                MessageBox.Show(TextHelper.CompanyMailMessage(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Check if User has successfully detected a Fake Email from ICT SIBE
        /// </summary>
        /// <param name="phishEmail"></param>
        private bool HandleFakeMail(MailItem phishEmail)
        {
            MailItem reportEmail = Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);

            if (IsFakeMail(phishEmail.HTMLBody))
            {

                // Read user data
                string userAccountName = Globals.ThisAddIn.GetUserIdentification();
                string computerName = Environment.MachineName;

                // Send report mail to SIBE
                string simSubject = $"SimPhish; {userAccountName}; {computerName}";
                simSubject = simSubject.Replace(",", ";");
                string body = $"Simulated Phishing attack reported by: {userAccountName}; {computerName};";
                body = body.Replace(",", ";");
                reportEmail.Subject = simSubject;
                reportEmail.To = PhishReporterConfig.FakeEmailAlias;
                reportEmail.Body = body;
                reportEmail.Send();
                
                TryDeleteEmailPermanentyFromFolder(phishEmail);

                // Show asyc messagebox before to avoid long waiting times
                MessageBox.Show(TextHelper.FakeMessageText(lang), Resources.MessageBox_Title, MessageBoxButtons.OK);

                // Delete Mail from Sent & Deleted folder
                DeleteEmailFromFolder(simSubject, OlDefaultFolders.olFolderSentMail);
                DeleteEmailFromFolder(simSubject, OlDefaultFolders.olFolderDeletedItems);

                DeleteEmailFromFolder("phishing", OlDefaultFolders.olFolderDeletedItems);

                return true;
            }

            return false;
        }

        /// <summary>
        /// Deletes the email from which the subject is passed in within the specified folder.
        /// Attention: The subject mustn't contain any special chars.
        /// </summary>
        private void DeleteEmailFromFolder(string subject, OlDefaultFolders folder)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNs = app.Application.GetNamespace("MAPI");
            MAPIFolder emailFolder = outlookNs.GetDefaultFolder(folder);

            Items restrictedItems = emailFolder.Items.Restrict("[Subject] = '" + subject + "'");
            int count = restrictedItems.Count;

            for (int i = count; i > 0; i--)
            {
                var mailItem = restrictedItems[i] as MailItem;

                if (mailItem == null)
                {
                    continue;
                }

                if (mailItem.Subject == subject)
                {
                    mailItem.Delete();
                    // Maybe just continue instead of break
                    // break;
                }

                Marshal.ReleaseComObject(mailItem);
            }
        }

        /// <summary>
        /// Deletes the email from which the subject is passed in within the specified folder.
        /// Attention: The subject mustn't contain any special chars.
        /// </summary>
        private void TryDeleteEmailPermanentyFromFolder(MailItem mailItem)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNs = app.Application.GetNamespace("MAPI");

            mailItem.Subject = "phishing";
            mailItem.Move(outlookNs.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems));

            mailItem.Delete();

            Marshal.ReleaseComObject(mailItem);
        }


        /// <summary>
        /// Check if Email is sent by known list of Fake senders
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
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
