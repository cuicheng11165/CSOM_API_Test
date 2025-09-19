using System.Net.Http.Headers;
using System.Net;

class RestApibase
{
    public async Task SendPostRequestAsync(string url, string token, HttpContent content)
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
