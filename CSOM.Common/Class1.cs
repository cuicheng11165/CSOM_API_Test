namespace CSOM.Common
{
    public class EnvConfig
    {
        static EnvConfig()
        {
            HouseName = File.ReadAllText("HouseName.txt");
            Authurization = File.ReadAllText("Authurization.txt");
        }

        public static string HouseName { get; set; }


        public static String Authurization { set; get; }

        public string GetSiteUrl(string siteRelativeUrl)
        {
            return $"https://{HouseName}.sharepoint.com/{siteRelativeUrl.TrimStart('/')}";
        }

    }
}
