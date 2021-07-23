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

            GetIndexProperty();
            return;
            System.Net.ServicePointManager.ServerCertificateValidationCallback = (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) => true;
            ClientContext context = new ClientContext("https://bigegg.sharepoint.com/sites/site202107231113");

            var userName = System.IO.File.ReadAllText(@"..\..\..\UserName\username.txt");
            var password = System.IO.File.ReadAllText(@"..\..\..\UserName\password.txt");


            SecureString se = new SecureString();
            foreach (var cc in password)
            {
                se.AppendChar(cc);
            }

            context.Credentials = new SharePointOnlineCredentials(userName, se);


            var web = context.Web;
            context.Load(web.AllProperties);
            context.ExecuteQuery();
            web.AllProperties["Index1"] = "Test";
            web.Update();
            context.ExecuteQuery();
            web.AddIndexedPropertyBagKey("Index1");

            context.ExecuteQuery();


        }


        private static void GetIndexProperty()
        {
            var indexPropertyKeys = "UAB1AGIAbABpAHMAaAAgAHQAbwAgAEQAaQByAGUAYwB0AG8AcgB5AA==|SQBuAGQAZQB4AFQAZQBzAHQA|RwBBAF8AUABvAGwAaQBjAHkARABpAHMAcABsAGEAeQBOAGEAbQBlAA==|RwBBAF8AUAByAGkAbQBhAHIAeQBTAGkAdABlAEMAbwBsAGwAZQBjAHQAaQBvAG4AQwBvAG4AdABhAGMAdAA=|RwBBAF8AUwBlAGMAbwBuAGQAYQByAHkAUwBpAHQAZQBDAG8AbABsAGUAYwB0AGkAbwBuAEMAbwBuAHQAYQBjAHQA|SQBuAGQAZQB4ADEA|";

            var keys = indexPropertyKeys.Split('|');


            foreach(var key in keys)
            {
                var value = Encoding.Unicode.GetString(Convert.FromBase64String(key));
            }


        }
    }
}
