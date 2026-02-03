using System;
using CSOM.Common;
using Microsoft.SharePoint.Client;

namespace UnifiedCsomTests.Scenarios
{
    internal static class SPListApiScenarios
    {
        internal static void PrintSiteTitle()
        {
            var token = EnvConfig.GetCsomToken();

            var siteRelativeUrl = "/contentstorage/CSP_f546b571-77f3-46c7-b8cb-e491cf3e3280";
            siteRelativeUrl = "/sites/site202503311557";

            var siteUrl = EnvConfig.GetSiteUrl(siteRelativeUrl);

            using ClientContext context = new ClientContext(siteUrl);

            context.ExecutingWebRequest += (object? sender, WebRequestEventArgs e) => { e.WebRequestExecutor.WebRequest.Headers[System.Net.HttpRequestHeader.Authorization] = token; };

            context.Load(context.Web);
            context.ExecuteQuery();

            Console.WriteLine(context.Web.Title);
        }


        internal static void UpdateListParser()
        {
            var token = EnvConfig.GetCsomToken();
            var siteRelativeUrl = "/sites/site202503311557";

            var siteUrl = EnvConfig.GetSiteUrl(siteRelativeUrl);

            using ClientContext context = new ClientContext(siteUrl);

            context.ExecutingWebRequest += (object? sender, WebRequestEventArgs e) => { e.WebRequestExecutor.WebRequest.Headers[System.Net.HttpRequestHeader.Authorization] = token; };

            var list = context.Web.Lists.GetByTitle("Documents");

            context.Load(list);
            context.ExecuteQuery();

            list.ParserDisabled = false;
            list.Update();
            context.ExecuteQuery();
        }
    }
}