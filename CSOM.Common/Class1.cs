using System.Runtime.CompilerServices;

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


        public static string GetSiteUrl(string siteRelativeUrl)
        {
            return $"https://{HouseName}.sharepoint.com/{siteRelativeUrl.TrimStart('/')}";
        }

        public static string GetAdminCenterUrl()
        {
            return $"https://{HouseName}-admin.sharepoint.com";
        }

        public static string GetToken()
        {
            return Authurization;
        }
    }
}
