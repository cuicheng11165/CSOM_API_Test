using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Security;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace CSOM_Authenticattion
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) => true;
            ClientContext context = new ClientContext("");

            SecureString se = new SecureString();
            foreach (var cc in "")
            {
                se.AppendChar(cc);
            }

            context.Credentials = new SharePointOnlineCredentials("", se);

            context.ExecutingWebRequest += context_ExecutingWebRequest;

            context.ExecuteQuery();

        }

        static void context_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            e.WebRequestExecutor.WebRequest.Proxy = new System.Net.WebProxy("10.2.6.69:8888");
        }
    }
}
