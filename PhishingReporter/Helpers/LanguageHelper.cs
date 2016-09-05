namespace PhishingReporter.Helpers
{
    using Microsoft.Win32;

    public class LanguageHelper
    {
        /// <summary>
        /// Try to read the language Key from LocalMachine or LocalUser Registry tree.
        /// Default language is English 1033
        /// </summary>
        /// <returns>Language enum (English as default or German)</returns>
        public static Language GetLanguage()
        {
            int languageCode = 1033; //Default to english 1031 for German
            int germanLanguageCode = 1031;

            string keyEntry = "UILanguage";
            string reg = @"Software\Microsoft\Office\14.0\Common\LanguageResources";

            try
            {
                RegistryKey k = Registry.CurrentUser.OpenSubKey(reg);
                if (k != null && k.GetValue(keyEntry) != null)
                {
                    languageCode = (int)k.GetValue(keyEntry);
                }

            }
            catch { }

            try
            {
                RegistryKey k = Registry.LocalMachine.OpenSubKey(reg);
                if (k != null && k.GetValue(keyEntry) != null)
                {
                    languageCode = (int)k.GetValue(keyEntry);
                }
            }
            catch { }

            return languageCode == germanLanguageCode ? Language.German : Language.English;
        }
    }
}