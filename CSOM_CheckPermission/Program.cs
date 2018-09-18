using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace CSOM_CheckPermission
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext context = new ClientContext("https://bigapp.sharepoint.com/sites/simmon1456");

            SecureString passworSecureString = new SecureString();

            System.IO.File.ReadAllText("password.txt").ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("simmon@baron.space", passworSecureString);

            var basePermission = context.Web.GetUserEffectivePermissions("i:0#.f|membership|simmon@baron.space");
            context.ExecuteQuery();

            var contributorType = context.Web.RoleDefinitions.GetByType(RoleType.Contributor);

            context.Load(contributorType);
            context.ExecuteQuery();

            var checkHigh = basePermission.Value.HasPermissions(48, 134287360);


        }
    }
}
