using System.Net.Http;
using System.Text.Json;
using CSOM.Common;


async Task<string> GetFoundryToken()
{
    var client = new HttpClient();

    string tenantId = EnvConfig.TenantId;
    string clientId = EnvConfig.ClientId;
    string clientSecret = "";

    var requestContent = new FormUrlEncodedContent(new[]
    {
        new KeyValuePair<string, string>("client_id", clientId),
        new KeyValuePair<string, string>("client_secret", clientSecret),
        new KeyValuePair<string, string>("scope", "https://ai.azure.com/.default"),
        new KeyValuePair<string, string>("grant_type", "client_credentials")
    });

    var response = await client.PostAsync(
        $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token",
        requestContent
    );

    var jsonString = await response.Content.ReadAsStringAsync();
    using var doc = JsonDocument.Parse(jsonString);

    // Return the access_token string
    return doc.RootElement.GetProperty("access_token").GetString();
}

var accessToken = await GetFoundryToken();

string configPath = Path.Combine("..", "..", "..", "..", "Config", "ClientSecretToken.txt");
System.IO.File.WriteAllText(configPath, accessToken);

Console.WriteLine("Successfully acquired AI Foundry token.");