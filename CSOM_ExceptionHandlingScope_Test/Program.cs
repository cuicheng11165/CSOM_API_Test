using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using Microsoft.SharePoint.Client;

namespace CSOM_ExceptionHandlingScope_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, errors) => true;

            using (ClientContext clientContext = new ClientContext("https://bigapp.sharepoint.com/sites/simmon1750"))
            {
               

                var currentWeb = clientContext.Web;

                var listGetById = currentWeb.Lists.GetByTitle("Documents");

                var folderUrl = "Shared%20Documents/f2";

                var exceptionHandlingScope = new ExceptionHandlingScope(clientContext);

                //List list = null;
                using (var currentScope = exceptionHandlingScope.StartScope())
                {
                    using (exceptionHandlingScope.StartTry())
                    {
                     
                        var folder = listGetById.RootFolder.Folders.GetByUrl(folderUrl);
                        folder.ListItemAllFields.BreakRoleInheritance(true,true);
                    }
                    using (exceptionHandlingScope.StartCatch())
                    {                    
                        var folder = listGetById.RootFolder.Folders.Add(folderUrl);
                        folder.ListItemAllFields.BreakRoleInheritance(true,true);
                    }
                }

                clientContext.ExecuteQuery();

                //Server端是否出现了异常
                Console.WriteLine("Server has Exception:" + exceptionHandlingScope.HasException);
                //Server端异常信息
                Console.WriteLine("Server Error Message:" + exceptionHandlingScope.ErrorMessage);



            }
        }



    }



}
