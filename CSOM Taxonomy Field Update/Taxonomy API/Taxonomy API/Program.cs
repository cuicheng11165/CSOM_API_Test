using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Taxonomy_API
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext context = new ClientContext("https://wrapperdev1102.sharepoint.com/sites/qlluo_Test1");

            SecureString se = new SecureString();
            foreach (var cc in "demo12!@")
            {
                se.AppendChar(cc);
            }

            context.Credentials = new SharePointOnlineCredentials("qlluo@wrapperdev1102.onmicrosoft.com", se);

            TaxonomySession session = TaxonomySession.GetTaxonomySession(context);

            context.Load(session.TermStores);
            context.ExecuteQuery();


            var termStore = session.TermStores[0];

            var termset = termStore.GetTermSet(new Guid("{95601aae-bed3-79dc-2591-562fa5d527f6}"));




            context.Load(termset.Terms);
            context.ExecuteQuery();




            CreateGroup(context, termStore);


            var termSets = termStore.GetTermSetsByName("TestTermSet1", 1033);

            context.Load(termSets);
            context.ExecuteQuery();


            context.Load(termSets[0].Terms);
            context.ExecuteQuery();

            var set0 = termSets[0];


            var term = set0.CreateTerm("testTerm1", 1033, Guid.NewGuid());

            var subTerm1 = term.CreateTerm("subTerm1", 1033, Guid.NewGuid());

            termStore.CommitAll();

            context.ExecuteQuery();



        }

        private static void CreateGroup(ClientContext context, TermStore termStore)
        {

            var group = termStore.CreateGroup("TestGroup1", Guid.NewGuid());

            var termSet = group.CreateTermSet("TestTermSet1", Guid.NewGuid(), 1033);

            var term = termSet.CreateTerm("testTerm1", 1033, Guid.NewGuid());

            var subTerm1 = term.CreateTerm("subTerm1", 1033, Guid.NewGuid());

            //termStore.CommitAll();

            context.ExecuteQuery();
        }
    }
}
