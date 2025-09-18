using System.Net;
using System.Runtime.CompilerServices;

namespace CSOM.Common
{
    public class EnvConfig
    {
        static EnvConfig()
        {
            HostName = File.ReadAllText(@"..\..\..\..\Config\HostName.txt");
            Authorization = File.ReadAllText(@"..\..\..\..\Config\Authorization.txt");
            ClientId = File.ReadAllText(@"..\..\..\..\Config\ClientId.txt");
            TenantId = File.ReadAllText(@"..\..\..\..\Config\TenantId.txt");
            UserName = File.ReadAllText(@"..\..\..\..\Config\UserName.txt");
            Password = File.ReadAllText(@"..\..\..\..\Config\Password.txt");
            CertificateThumbprint = File.ReadAllText(@"..\..\..\..\Config\CertificateThumbprint.txt");
        }

        public static string HostName { get; set; }

        public static String Authorization { set; get; }

        public static String ClientId { set; get; }

        public static String TenantId { set; get; }
        public static String UserName { set; get; }
        public static String Password { set; get; }

        public static String CertificateThumbprint { set; get; }


        public static string GetSiteUrl(string siteRelativeUrl)
        {
            return $"https://{HostName}/{siteRelativeUrl.TrimStart('/')}";
        }

        public static string GetAdminCenterUrl()
        {
            return "https://" + HostName.Replace(".sharepoint.com", "-admin.sharepoint.com");
        }

        public static string GetToken()
        {
            return Authorization;
        }
    }

}