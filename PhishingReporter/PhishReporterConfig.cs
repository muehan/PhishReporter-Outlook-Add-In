namespace PhishingReporter
{
    public class PhishReporterConfig
    {
        // Report Configuration
        public static string SecurityTeamEmailAlias = "servicedesk@your-company.com";

        public static string ReportEmailSubject = "Phishing[phish:Secure]";

        // Ribbon Group Config
        public static string RibbonGroupName = "ICT Security";

        // Button Config
        public static string ButtonName = "Report Phishing";
        public static string ButtonHoverDescription =
            "Report a suspicious email to the Company Information Security Team.";

        public static string ButtonScreenTip = "Report phishing emails";

        public static string ButtonSuperTip =
            "Use this button to report suspicious emails to the Company Information Security team.";

        public static string FakePhishingUrl =
            "fake-url1.net, fake-url2.com, fake-url3.org";

    }
}
