using System.Runtime.CompilerServices;

namespace CSOM.Common
{
    public class EnvConfig
    {
        static EnvConfig()
        {
            HostName = File.ReadAllText(@"..\..\..\..\HostName.txt");
            Authurization = File.ReadAllText(@"..\..\..\..\Authurization.txt");
        }

        public static string HostName { get; set; }

        public static String Authurization { set; get; }


        public static string GetSiteUrl(string siteRelativeUrl)
        {
            return $"https://{HostName}/{siteRelativeUrl.TrimStart('/')}";
        }

        public static string GetAdminCenterUrl()
        {
            return "https://"+ HostName.Replace(".sharepoint.com", "-admin.sharepoint.com");    
        }

        public static string GetToken()
        {
            return Authurization;
        }
    }
}
