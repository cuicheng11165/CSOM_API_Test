using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using CSOM.Common;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;

namespace ExportCsomTokenTest
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string siteUrl = EnvConfig.GetSiteUrl("/sites/site202503311557"); ;
            string tenantId = EnvConfig.TenantId;
            string clientId = EnvConfig.ClientId;
            string certificateThumbprint = EnvConfig.CertificateThumbprint;
            
            string[] scopes = new[] { $"https://{new Uri(siteUrl).Host}/.default" };

            // Find certificate by thumbprint in the local machine store
            X509Certificate2 certificate = FindCertificateByThumbprint(certificateThumbprint);
            if (certificate == null)
            {
                Console.WriteLine("Certificate not found.");
                return;
            }

            // --- Use MSAL to get access token ---
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithCertificate(certificate)
                .Build();

            AuthenticationResult authResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            string accessToken = authResult.AccessToken;


            System.IO.File.WriteAllText("..\\..\\..\\..\\Authorization.txt", "Bearer " + accessToken);

            // --- Connect with CSOM using the access token ---
            using (var context = new ClientContext(siteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                };

                context.Load(context.Web, w => w.Title);
                await context.ExecuteQueryAsync();

                Console.WriteLine($"Connected to: {context.Web.Title}");
            }
            Console.ReadLine();
        }

        private static X509Certificate2 FindCertificateByThumbprint(string thumbprint)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certs.Count > 0)
                {
                    return certs[0];
                }
            }
            return null;
        }
    }
}
