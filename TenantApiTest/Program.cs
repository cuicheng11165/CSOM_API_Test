using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace TenantApiTest
{
    class Program
    {
        static void Main(string[] args)
        {
            SecureString se = new SecureString();
            foreach (var cc in "1DC")
            {
                se.AppendChar(cc);
            }

            ClientContext context = new ClientContext("https://-admin.sharepoint.com");
            context.Credentials = new SharePointOnlineCredentials("@.space", se);

            Tenant tenant = new Tenant(context);

            var hubSites=tenant.GetHubSitesProperties();
            var sites=tenant.GetKnowledgeHubSite();


            var hubSiteProperties=tenant.GetHubSitesProperties();



                     
            context.Load(hubSiteProperties);
            context.Load(hubSites);
            context.ExecuteQuery();

            foreach(var site in hubSiteProperties)
            {
                var url=site.SiteUrl;
                var property=tenant.GrantHubSiteRights(url, new[] { "DL_GA_DEV1" }, SPOHubSiteUserRights.Join);
                context.ExecuteQuery();
            }
        }
    }
}
