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
            foreach (var cc in "Avepoint@2025Q2")
            {
                se.AppendChar(cc);
            }

            ClientContext context = new ClientContext("https://cloudgov.sharepoint.com");
            context.Credentials = new SharePointOnlineCredentials("simmon@cloudgov.onmicrosoft.com", se);

            context.Load(context.Web);
            context.ExecuteQuery();
           
            Tenant tenant = new Tenant(context); 

            var containerTypes1 = tenant.GetSPOContainerTypes(SPContainerTypeTenantType.OwningTenant);
            context.ExecuteQuery();

            var containerTypes2 = tenant.GetSPOContainerTypes(SPContainerTypeTenantType.ConsumingTenant);
            context.ExecuteQuery();


            var containers = tenant.GetSPOContainersByApplicationId(new Guid("a187e399-0c36-4b98-8f04-1edc167a0996"), false, "");

            context.ExecuteQuery();

            foreach (var c in containers.Value.ContainerCollection)
            {
                var re = tenant.GetSPOContainerByContainerId(c.ContainerId);

                context.ExecuteQuery();
            }




            var url = "https://bigapp.sharepoint.com/sites/waltTestCreateSiteCollection2";


            var siteProperties = tenant.GetSitePropertiesByUrl(url, false);
            context.Load(siteProperties);
            context.ExecuteQuery();

            siteProperties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
            siteProperties.Update();

            //var hubSiteProperties=tenant.GetHubSitesProperties();




            //context.Load(hubSiteProperties);

            //context.Load(hubSites);
            context.ExecuteQuery();

            //foreach(var site in hubSiteProperties)
            //{
            //    var url=site.SiteUrl;
            //    var property=tenant.GrantHubSiteRights(url, new[] { "DL_GA_DEV1" }, SPOHubSiteUserRights.Join);
            //    context.ExecuteQuery();
            //}
        }

        private static void Context_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
