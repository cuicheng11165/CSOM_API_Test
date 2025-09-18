using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Security;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace UpdateWebAllProperties
{
    class Program
    {
        static void Main(string[] args)
        {


            System.Net.ServicePointManager.ServerCertificateValidationCallback = (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) => true;
            ClientContext context = new ClientContext("https://wfmrm.sharepoint.com/sites/Eira_Group_CA_03");

            var userName = "eiraadmin@wfmrm.onmicrosoft.com";
            var password = "QDFasd@12333";
            SecureString se = new SecureString();
            foreach (var cc in password)
            {
                se.AppendChar(cc);
            }


            ClientContext context1 = new ClientContext("https://wfmrm-admin.sharepoint.com");
            context1.Credentials = new SharePointOnlineCredentials(userName, se);
            var tenant = new Tenant(context1);
            var siteProperties = tenant.GetSitePropertiesByUrl("https://wfmrm.sharepoint.com/sites/Eira_Group_CA_03", false);
            siteProperties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
            siteProperties.Update();
            context1.ExecuteQuery();

            context.Credentials = new SharePointOnlineCredentials(userName, se);


            var web = context.Web;
            context.Load(web.AllProperties);
            context.ExecuteQuery();
            web.AllProperties["Eira_Radio_01"] = null;
            web.Update();
            context.ExecuteQuery();
        


        }


        private static void GetIndexProperty()
        {
            var indexPropertyKeys = "UAB1AGIAbABpAHMAaAAgAHQAbwAgAEQAaQByAGUAYwB0AG8AcgB5AA==|SQBuAGQAZQB4AFQAZQBzAHQA|RwBBAF8AUABvAGwAaQBjAHkARABpAHMAcABsAGEAeQBOAGEAbQBlAA==|RwBBAF8AUAByAGkAbQBhAHIAeQBTAGkAdABlAEMAbwBsAGwAZQBjAHQAaQBvAG4AQwBvAG4AdABhAGMAdAA=|RwBBAF8AUwBlAGMAbwBuAGQAYQByAHkAUwBpAHQAZQBDAG8AbABsAGUAYwB0AGkAbwBuAEMAbwBuAHQAYQBjAHQA|SQBuAGQAZQB4ADEA|";

            var keys = indexPropertyKeys.Split('|');


            foreach (var key in keys)
            {
                var value = Encoding.Unicode.GetString(Convert.FromBase64String(key));
            }


        }
    }
}
