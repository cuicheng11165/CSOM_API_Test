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
            foreach (var cc in " # ")
            {
                se.AppendChar(cc);
            }

            ClientContext context = new ClientContext("https://d-admin.sharepoint.com");
            context.Credentials = new SharePointOnlineCredentials("simmon@ .space", se);

            Tenant tenant = new Tenant(context);

            var sites=tenant.GetDeletedSiteProperties(0);
            
 
            
            context.Load(sites);
            context.ExecuteQuery();
        }
    }
}
