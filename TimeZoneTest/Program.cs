using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;

namespace TimeZoneTest
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            TestServerAPI();
            TestClientAPI();
        }

        private static void TestClientAPI()
        {
            Console.WriteLine("Test Client API");
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

        private static void TestServerAPI()
        {
            Console.WriteLine("Test Server API");
            using (SPSite site = new SPSite("http://win-cpqm71buqvj:1000/sites/Test"))
            {
                var file = site.RootWeb.GetFile("/sites/Test/DocumentTest/12345.txt");

                var listItem = file.Item;

                DateTime dateTime = new DateTime(2014, 1, 10, 0, 0, 0, 0, DateTimeKind.Local);

                ServerAPISetModified(listItem, dateTime);
                ServerAPIOutputModified(listItem);

                //Time: 2015/1/1 0:00:00 , Kind: Unspecified

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

        private static void ClientAPIOutputModified(ClientContext context, ListItem listItem)
        {
            context.Load(listItem);
            context.ExecuteQuery();
            var modifiedTime = (DateTime) listItem["Modified"];

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
      
        private static void ServerAPISetModified(SPListItem listItem, DateTime dateTime)
        {
            Console.WriteLine("Set Time: {0} , Kind: {1}", dateTime, dateTime.Kind);
            listItem["Modified"] = dateTime;
            listItem.Update();
        }

        private static void ServerAPIOutputModified(SPListItem listItem)
        {
            var modifiedTime = (DateTime) listItem["Modified"];

            Console.WriteLine("Time: {0} , Kind: {1}", modifiedTime, modifiedTime.Kind);
            Console.WriteLine("\r\n");
        }
    }
}