using System;
using Microsoft.SharePoint.Client;

#if !NET8_0_OR_GREATER
using Microsoft.SharePoint;
#endif

namespace TimeZoneTest
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            TestClientAPI_Document();
            TestClientAPI_ListItem();
            TestServerAPI_Document();
            TestServerAPI_ListItem();
        }

        private static void TestClientAPI_Document()
        {
            Console.WriteLine("Test Client API for Document");
            ClientContext context = new ClientContext("http://win-cpqm71buqvj:1000/sites/Test");

            var web = context.Web;
            var file = web.GetFileByServerRelativeUrl("/sites/Test/DocumentTest/12345.txt");
            var listItem = file.ListItemAllFields;

            DateTime dateTime = new DateTime(2014, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
            ClientAPISetModified(context, listItem, dateTime);
            ClientAPIOutputModified(context, listItem);

            DateTime dateTime1 = new DateTime(2015, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
            ClientAPISetModified(context, listItem, dateTime1);
            ClientAPIOutputModified(context, listItem);

            DateTime dateTime2 = new DateTime(2016, 1, 10, 0, 0, 0, 0, DateTimeKind.Utc);
            ClientAPISetModified(context, listItem, dateTime2);
            ClientAPIOutputModified(context, listItem);

            DateTime dateTime3 = new DateTime(2017, 1, 10, 0, 0, 0, 0, DateTimeKind.Unspecified);
            ClientAPISetModified(context, listItem, dateTime3);
            ClientAPIOutputModified(context, listItem);
        }

        private static void TestClientAPI_ListItem()
        {
            Console.WriteLine("Test Client API for ListItem");
            ClientContext context = new ClientContext("http://win-cpqm71buqvj:1000/sites/Test");

            var web = context.Web;
            var list = web.Lists.GetByTitle("customList");
            var listItem = list.GetItemById(1);

            DateTime dateTime = new DateTime(2014, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
            ClientAPISetModified(context, listItem, dateTime);
            ClientAPIOutputModified(context, listItem);

            DateTime dateTime1 = new DateTime(2015, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
            ClientAPISetModified(context, listItem, dateTime1);
            ClientAPIOutputModified(context, listItem);

            DateTime dateTime2 = new DateTime(2016, 1, 10, 0, 0, 0, 0, DateTimeKind.Utc);
            ClientAPISetModified(context, listItem, dateTime2);
            ClientAPIOutputModified(context, listItem);

            DateTime dateTime3 = new DateTime(2017, 1, 10, 0, 0, 0, 0, DateTimeKind.Unspecified);
            ClientAPISetModified(context, listItem, dateTime3);
            ClientAPIOutputModified(context, listItem);
        }

#if NET8_0_OR_GREATER
        private static void TestServerAPI_ListItem()
        {
            Console.WriteLine("Server API sample is not supported when targeting .NET 8.");
        }

        private static void TestServerAPI_Document()
        {
            Console.WriteLine("Server API sample is not supported when targeting .NET 8.");
        }
#else
        private static void TestServerAPI_ListItem()
        {
            Console.WriteLine("Test Server API for ListItem");
            using (SPSite site = new SPSite("http://win-cpqm71buqvj:1000/sites/Test"))
            {
                var list = site.RootWeb.Lists["customList"];
                var listItem = list.GetItemById(1);

                DateTime dateTime = new DateTime(2014, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
                ServerAPISetModified(listItem, dateTime);
                ServerAPIOutputModified(listItem);

                DateTime dateTime1 = new DateTime(2015, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
                ServerAPISetModified(listItem, dateTime1);
                ServerAPIOutputModified(listItem);

                DateTime dateTime2 = new DateTime(2016, 1, 10, 0, 0, 0, 0, DateTimeKind.Utc);
                ServerAPISetModified(listItem, dateTime2);
                ServerAPIOutputModified(listItem);

                DateTime dateTime3 = new DateTime(2017, 1, 10, 0, 0, 0, 0, DateTimeKind.Unspecified);
                ServerAPISetModified(listItem, dateTime3);
                ServerAPIOutputModified(listItem);
            }
        }

        private static void TestServerAPI_Document()
        {
            Console.WriteLine("Test Server API for Document");
            using (SPSite site = new SPSite("http://win-cpqm71buqvj:1000/sites/Test"))
            {
                var file = site.RootWeb.GetFile("/sites/Test/DocumentTest/12345.txt");
                var listItem = file.Item;

                DateTime dateTime = new DateTime(2014, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
                ServerAPISetModified(listItem, dateTime);
                ServerAPIOutputModified(listItem);

                DateTime dateTime1 = new DateTime(2015, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);
                ServerAPISetModified(listItem, dateTime1);
                ServerAPIOutputModified(listItem);

                DateTime dateTime2 = new DateTime(2016, 1, 10, 0, 0, 0, 0, DateTimeKind.Utc);
                ServerAPISetModified(listItem, dateTime2);
                ServerAPIOutputModified(listItem);

                DateTime dateTime3 = new DateTime(2017, 1, 10, 0, 0, 0, 0, DateTimeKind.Unspecified);
                ServerAPISetModified(listItem, dateTime3);
                ServerAPIOutputModified(listItem);
            }
        }
#endif

        private static void ClientAPIOutputModified(ClientContext context, ListItem listItem)
        {
            context.Load(listItem);
            context.ExecuteQuery();
            var modifiedTime = (DateTime)listItem["Modified"];

            Console.WriteLine("Output Time: {0} , Kind: {1}", modifiedTime, modifiedTime.Kind);
            Console.WriteLine("\r\n");
        }

        private static void ClientAPISetModified(ClientContext context, ListItem listItem, DateTime dateTime)
        {
            Console.WriteLine("Set Time: {0} , Kind: {1}", dateTime, dateTime.Kind);

            listItem["Modified"] = dateTime;
            listItem.Update();
            context.ExecuteQuery();
        }

#if !NET8_0_OR_GREATER
        private static void ServerAPISetModified(SPListItem listItem, DateTime dateTime)
        {
            Console.WriteLine("Set Time: {0} , Kind: {1}", dateTime, dateTime.Kind);
            listItem["Modified"] = dateTime;
            listItem.Update();
        }

        private static void ServerAPIOutputModified(SPListItem listItem)
        {
            var modifiedTime = (DateTime)listItem["Modified"];

            Console.WriteLine("Output Time: {0} , Kind: {1}", modifiedTime, modifiedTime.Kind);
            Console.WriteLine("\r\n");
        }
#endif
    }
}
