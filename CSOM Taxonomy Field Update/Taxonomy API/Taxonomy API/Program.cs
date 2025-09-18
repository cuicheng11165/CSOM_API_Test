using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections;
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
            var context = new ClientContext("https://bigapp.sharepoint.com/sites/mmstest001");

         

      

            Create(context);



            TaxonomySession session = TaxonomySession.GetTaxonomySession(context);

            context.Load(session.TermStores);
            context.ExecuteQuery();


            var termStore = session.TermStores[0];


            var termset = termStore.GetTermSet(new Guid("a8432e52-4018-479f-acbf-d0e06150f656"));
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


        private static void Create(ClientContext context)
        {
            {

                // Load the web
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                // Get the Taxonomy Session and default Term Store
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
                context.Load(taxonomySession);
                context.ExecuteQuery();

                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                context.Load(termStore);
                context.ExecuteQuery();

                // Retrieve a specific TermSet (Replace with your TermSet GUID)

                Guid termSetId = new Guid("a8432e52-4018-479f-acbf-d0e06150f656");
                TermSet termSet = termStore.GetTermSet(termSetId);
                context.Load(termSet);
                context.ExecuteQuery();

                // Prepare the field name and ID
                string columnname = "MyManagedMetadataField" + Guid.NewGuid();
                Guid fieldId = Guid.NewGuid();

                // Prepare field XML. We use TaxonomyFieldType here.
                // Note that in this XML, we set placeholders for SspId, TermSetId, etc.
                Field f = web.Fields.AddFieldAsXml("<Field Type='TaxonomyFieldType'   Name='" + columnname + "' DisplayName='" + columnname + "'  ShowField='Term1033' />", false, AddFieldOptions.DefaultValue);

                context.Load(f);
                context.ExecuteQuery();

                TaxonomyField taxField = context.CastTo<TaxonomyField>(f);

                taxField.SspId = new Guid("1b7c6a2a-e692-4cef-baee-ba089fccba51");
                taxField.TermSetId = new Guid("a8432e52-4018-479f-acbf-d0e06150f656");
                taxField.AllowMultipleValues = false;
                taxField.Open = true;

                taxField.TargetTemplate = string.Empty;
                taxField.AnchorId = Guid.Empty;
                taxField.Update();
                //list.Update();
                context.ExecuteQuery();


                var newField = web.Fields.GetById(f.Id);
                context.Load(newField);
                context.ExecuteQuery();


                if (newField.TypeAsString != f.TypeAsString)
                {
                    throw new Exception();
                }






                // Cast to TaxonomyField to set Taxonomy-specific properties
                //TaxonomyField taxField = context.CastTo<TaxonomyField>(newField);
                //taxField.SspId = termStore.Id;
                //taxField.TermSetId = termSet.Id;
                //taxField.AnchorId = Guid.Empty; // Use Guid.Empty if no anchor term is needed
                //taxField.Update();
                //context.ExecuteQuery();
            }
        }
    }
}
