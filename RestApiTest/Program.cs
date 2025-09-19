using System.Text.Json;
using CSOM.Common;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;

public class UserRoleInfo
{
    public string? ContainerId { get; set; }
    public string? loginName { get; set; }
    public string? role { get; set; }

    public override string ToString()
    {
        return JsonSerializer.Serialize(this);
    }
}

class RestApi
{


    static void Main()
    {

        AddSPOContainerUserRole();

    }

    public static void AddSPOContainerUserRole()
    {
        var adminUrl = EnvConfig.GetAdminCenterUrl();

        var token = EnvConfig.GetCsomToken();

        var url = $"{adminUrl}/_api/SPO.Tenant/AddSPOContainerUserRole";


        var body = new UserRoleInfo
        {
            ContainerId = "b!cbVG9fN3x0a4y-SRzz4ygBN1rEsDORVIoVoqpy-MK7iYCCQMZhJ6T5h-mZKqu459",
            loginName = "Eric@cloudgov.onmicrosoft.com",
            role = "owner"
        };

        SendPostRequestAsync(url, token, new StringContent(body.ToString())).GetAwaiter().GetResult();



    }


    public static void RemoveSPOContainerUserRole()
    {
        var adminUrl = EnvConfig.GetAdminCenterUrl();

        var token = EnvConfig.GetCsomToken();

        var url = $"{adminUrl}/_api/SPO.Tenant/RemoveSPOContainerUserRole";


        var body = new UserRoleInfo
        {
            ContainerId = "b!cbVG9fN3x0a4y-SRzz4ygBN1rEsDORVIoVoqpy-MK7iYCCQMZhJ6T5h-mZKqu459",
            loginName = "Eric@cloudgov.onmicrosoft.com",
            role = "owner"
        };

        SendPostRequestAsync(url, token, new StringContent(body.ToString())).GetAwaiter().GetResult();



    }

    public static async Task SendPostRequestAsync(string url, string token, HttpContent content)
    {
        // Create a proxy instance
        var proxy = new WebProxy("http://localhost:8888");
        // Create handler with proxy

        var handler = new HttpClientHandler
        {
            Proxy = proxy,
            UseProxy = true
        };
        using (var client = new HttpClient(handler))
        {

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
            
            var response = await client.PostAsync(url, content);
            string responseBody = await response.Content.ReadAsStringAsync();
            // Handle response as needed
        }
    }
}
