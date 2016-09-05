using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace PhishingReporter
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;

        private Outlook.MAPIFolder spamFolder;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = Application.Inspectors;
            spamFolder = this.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        public Outlook.MAPIFolder GetSpamFolder()
        {
            return spamFolder;
        }
    }
}
