namespace PhishingReporter.Helpers
{
    using System;

    using PhishingReporter.Properties;

    public static class TextHelper
    {
        public static string MessageBoxText(Language lang)
        {
            if (lang == Language.German)
            {
                return "Handelt es sich bei dieser E-Mail um Spam (z.B. Newsletter, Marketing-Kampagne, etc.)\n\r"
                     + "oder um eine Phishing-Attacke (z.B. Bekanntgabe von Usernamen, Passwörter,\n\r"
                     + "Kreditkarteninformationen, Zahlungsfreigaben etc.)?";
            }

            return "Is this a Spam email (e.g. newsletter, marketing campaign, etc.) or a Phishing-Attack\n\r"
                 + "(e.g. asking for username, passwords, credit card information, request for payment etc.)?";

        }

        public static string FakeMessageText(Language lang)
        {
            if (lang == Language.German)
            {
                return
                    "Glückwünsch! Bei diesem E-Mail handelte es sich um eine simulierte Phishing-Attacke von asdf AG. Sie haben richtig reagiert!\n\r\n\rBitte haben Sie etwas Geduld, bis das E-Mail gelöscht wird.";
            }

            return
                "Congratulations! The email you reported was a simulated phishing attack initiated by asdf Ltd. Good job!\n\r\n\rPlease be patient while the email is being deleted.";
        }

        public static string ThankYouMessage(Language lang)
        {
            if (lang == Language.German)
            {
                return Resources.ThankYouMsgBox_Text_german;
            }

            return Resources.ThankYouMsgBox_Text_english;
        }

        public static string NoEmailSelectedText(Language lang)
        {
            if (lang == Language.German)
            {
                return Resources.NoMessageText_german;
            }

            return Resources.NoMessageText_english;
        }

        public static string SpamAddedText(Language lang)
        {

            if (lang == Language.German)
            {
                return Resources.SpamText_german;
            }

            return Resources.SpamText_english;
        }

        public static string CompanyMailMessage(Language lang)
        {
            if (lang == Language.German)
            {
                return Resources.CompanyMail_german;
            }

            return Resources.CompanyMail_english;
        }
    }
}
