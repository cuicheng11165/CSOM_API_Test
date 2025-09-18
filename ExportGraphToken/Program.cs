using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

class Program
{
    static async Task Main(string[] args)
    {
        // --- Configuration ---
        string tenantId = "7ce674d3-a55a-43ba-806b-0da0801a9a6a";
        string clientId = "dcd331a7-9462-4a88-a2ca-5a2c785c1cf1";
        string certificateThumbprint = "6a032348581f617842f29b3f45a385e382d5b1e3";

        // --- Define the scope for the Graph API ---
        // For app-only authentication, use the /.default scope.
        string[] scopes = new[] { "https://graph.microsoft.com/.default" };

        // --- Get the certificate from the store ---
        X509Certificate2 certificate = FindCertificateByThumbprint(certificateThumbprint);
        if (certificate == null)
        {
            Console.WriteLine("Certificate not found.");
            return;
        }

        try
        {
            // --- Build MSAL Confidential Client Application ---
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithCertificate(certificate)
                .Build();

            // --- Acquire the token silently ---
            AuthenticationResult authResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            string accessToken = authResult.AccessToken;

            Console.WriteLine("Successfully acquired Graph API token.");

            // --- Use the token to call the Graph API ---
            await CallGraphApi(accessToken);
        }
        catch (MsalException msalex)
        {
            Console.WriteLine($"Error acquiring token: {msalex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    private static X509Certificate2 FindCertificateByThumbprint(string thumbprint)
    {
        // Search the CurrentUser certificate store
        using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);
            var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
            if (certs.Count > 0)
            {
                return certs[0];
            }
        }
        // Search the LocalMachine certificate store
        using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
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

    private static async Task CallGraphApi(string accessToken)
    {
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Example: Get a list of users
            string graphEndpoint = "https://graph.microsoft.com/v1.0/users";
            var response = await httpClient.GetAsync(graphEndpoint);

            if (response.IsSuccessStatusCode)
            {
                string jsonResponse = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Graph API call successful. Response: {jsonResponse}");
            }
            else
            {
                Console.WriteLine($"Graph API call failed with status code: {response.StatusCode}");
                string errorResponse = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Error details: {errorResponse}");
            }
        }
    }
}
