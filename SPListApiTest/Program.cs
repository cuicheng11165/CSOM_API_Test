using CSOM.Common;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;

var token = EnvConfig.GetToken();
var siteRelativeUrl = "/contentstorage/CSP_f546b571-77f3-46c7-b8cb-e491cf3e3280";

var siteUrl = EnvConfig.GetSiteUrl(siteRelativeUrl);



ClientContext context = new ClientContext(siteUrl);

context.ExecutingWebRequest += (object? sender, WebRequestEventArgs e) =>
{
    e.WebRequestExecutor.WebRequest.Headers[System.Net.HttpRequestHeader.Authorization] = token;
};


context.Load(context.Web, w=>w.Url);
context.ExecuteQuery();


Console.WriteLine(context.Web.Title);
