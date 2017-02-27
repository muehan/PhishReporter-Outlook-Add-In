namespace PhishingReporter
{
    public class PhishReporterConfig
    {
        // Report Configuration
        public static string SecurityTeamEmailAlias = "servicedesk@company.com";

        public static string FakeEmailAlias = "phishing@company.com";

        public static string ReportEmailSubject = "Phishing[phish:Secure]";

        // Ribbon Group Config
        public static string RibbonGroupName = "ICT Security";

        // Button Config
        public static string ButtonName = "Report Phishing";
        public static string ButtonHoverDescription =
            "Report a suspicious email to the company Information Security Team.";

        public static string ButtonScreenTip = "Report phishing emails";

        public static string ButtonSuperTip =
            "Use this button to report suspicious emails to the company Information Security team.";

        // create sructured file for this!
        public static string FakePhishingUrl =
                "url1.net,"
                + "url2.net,"
                + "more-url-here.com";
    }
}
