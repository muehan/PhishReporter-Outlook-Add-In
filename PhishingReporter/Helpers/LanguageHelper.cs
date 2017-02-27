namespace PhishingReporter.Helpers
{
    using Microsoft.Win32;

    public class LanguageHelper
    {
        /// <summary>
        /// Try to read the language Key from LocalMachine or LocalUser Registry tree.
        /// Default language is English 1033
        /// </summary>
        /// <returns>office language key</returns>
        public static Language ParseLanguage(int code)
        {
            int germanLanguageCode = 1031;
            
            return code == germanLanguageCode ? Language.German : Language.English;
        }
    }
}
