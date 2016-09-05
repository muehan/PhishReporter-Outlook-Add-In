using System;

namespace PhishingReporter.Helpers
{
    using PhishingReporter.Properties;

    public static class TextHelper
    {
        public static string MessageBoxText(Language lang)
        {
            if (lang == Language.German)
            {
                return "Diese markierte Nachricht wird weitergeleitet an "
                                   + PhishReporterConfig.SecurityTeamEmailAlias + Environment.NewLine
                                   + "und aus Ihrem Posteingang entfernt. Möchten Sie fortfahren?";
            }

            return "The selected message will be forwarded to "
                                 + PhishReporterConfig.SecurityTeamEmailAlias + Environment.NewLine
                                 + " and removed from your inbox. Would you like to continue?";

        }

        public static string FakeMessageText(Language lang)
        {
            if (lang == Language.German)
            {
                return Resources.FakeMessage_Text_german;
            }

            return Resources.FakeMessage_Text_english;
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
    }
}
