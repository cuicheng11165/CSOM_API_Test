using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;

namespace CamlQueryTest
{
    class Program
    {
        static void Main(string[] args)
        {
            CalmQuery();
        }

        private static void BasicQuery()
        {
            string siteUrl = "http://MyServer/sites/MySiteCollection";

            ClientContext clientContext = new ClientContext(siteUrl);
            List oList = clientContext.Web.Lists.GetByTitle("Announcements");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"
                                    <View>
                                      <Query>
                                        <Where>
                                          <Geq>
                                            <FieldRef Name='ID'/>
                                            <Value Type='Number'>10</Value>
                                          </Geq>
                                        </Where>
                                      </Query>
                                      <RowLimit>100</RowLimit>
                                    </View>";

            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem);

            clientContext.ExecuteQuery();

            foreach (ListItem oListItem in collListItem)
            {
                Console.WriteLine("ID: {0} \nTitle: {1} \nBody: {2}", oListItem.Id, oListItem["Title"], oListItem["Body"]);
            }
        }

        private static void CalmQuery()
        {
            string siteUrl = "http://win-cpqm71buqvj:1000/sites/Test";

            ClientContext clientContext = new ClientContext(siteUrl);
            List oList = clientContext.Web.Lists.GetByTitle("VersionTest");

            clientContext.Load(oList);
            clientContext.ExecuteQuery();

            ListItemCollection listItems = null;
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><RowLimit>2</RowLimit></View>";

            camlQuery.FolderServerRelativeUrl = "/sites/Test/VersionTest";
            ListItemCollectionPosition pos = null;

            do
            {
                camlQuery.ListItemCollectionPosition = pos;

                listItems = oList.GetItems(camlQuery);
                clientContext.Load(listItems, items => items.ListItemCollectionPosition, items => items.IncludeWithDefaultProperties(item => item["FSObjType"]).Where(item => (string)item["FSObjType"] == "0"));
                clientContext.ExecuteQuery();
                foreach (ListItem item in listItems)
                {


                }
                pos = listItems.ListItemCollectionPosition;
            }
            while (pos != null);

        }
    }
}
