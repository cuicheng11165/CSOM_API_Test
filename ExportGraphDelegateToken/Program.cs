using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.Security;
using CSOM.Common;

class Program
{
    static async Task Main()
    {
        // ====== CONFIG ======
        string tenantId = EnvConfig.TenantId;
        string clientId = EnvConfig.ClientId;
        string username = EnvConfig.UserName;
        string passwordPlain = EnvConfig.Password;


        string[] scopes = new[] { "User.Read" };             // Delegated scopes (NOT .default here)

        // ====== Build a public client for ROPC ======
        var app = PublicClientApplicationBuilder
            .Create(clientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, tenantId) // change cloud if GCC/China/Germany
            .WithRedirectUri("http://localhost")                     // not used by ROPC, but harmless
            .Build();

        // Convert to SecureString (MSAL requires it for ROPC)
        using var securePwd = new SecureString();
        foreach (var c in passwordPlain) securePwd.AppendChar(c);

        try
        {
            // ====== Acquire delegated token via username + password ======
            var result = await app
                .AcquireTokenByUsernamePassword(scopes, username, securePwd)
                .ExecuteAsync();

            Console.WriteLine("Access token acquired.");
            // Console.WriteLine(result.AccessToken); // uncomment if you want to see the raw token

            // ====== Use the token (call Graph /me) ======
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", result.AccessToken);

            var resp = await http.GetAsync("https://graph.microsoft.com/v1.0/me");
            var body = await resp.Content.ReadAsStringAsync();

            Console.WriteLine($"Status: {(int)resp.StatusCode} {resp.ReasonPhrase}");
            Console.WriteLine(body);
        }
        catch (MsalUiRequiredException ex)
        {
            Console.WriteLine("MSAL UI required (ROPC blocked by policy, CA, or MFA).");
            Console.WriteLine(ex.Message);
        }
        catch (MsalServiceException ex)
        {
            Console.WriteLine("MSAL service error (tenant/app/permissions/username issues).");
            Console.WriteLine($"ErrorCode: {ex.ErrorCode}");
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error:");
            Console.WriteLine(ex);
        }
    }
}
