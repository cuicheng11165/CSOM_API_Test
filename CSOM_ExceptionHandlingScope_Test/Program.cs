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

            using (ClientContext clientContext = new ClientContext("https://cnblogtest.sharepoint.com"))
            {
                var pasword = new SecureString();
                "abc123!@#".ToCharArray().ToList().ForEach(pasword.AppendChar);

                clientContext.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", pasword);//设置权限

                var currentWeb = clientContext.Web;


                var exceptionHandlingScope = new ExceptionHandlingScope(clientContext);

                //List list = null;
                using (var currentScope = exceptionHandlingScope.StartScope())
                {
                    using (exceptionHandlingScope.StartTry())
                    {
                        //此API调用时，如果此List在Server端不存在，会出现异常。
                        var listGetById = currentWeb.Lists.GetByTitle("Documents Test");
                        listGetById.Description = "List Get By Id";
                        listGetById.Update();
                    }
                    using (exceptionHandlingScope.StartCatch())
                    {
                        ListCreationInformation listCreationInfo = new ListCreationInformation();
                        listCreationInfo.Title = "Documents Test";
                        listCreationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
                        listCreationInfo.Description = "List create in catch block";
                        currentWeb.Lists.Add(listCreationInfo);
                    }
                }

                List list = currentWeb.Lists.GetByTitle("Documents Test");
                clientContext.Load(list);
                clientContext.ExecuteQuery();//执行查询,不会出异常

                //Server端是否出现了异常
                Console.WriteLine("Server has Exception:" + exceptionHandlingScope.HasException);
                //Server端异常信息
                Console.WriteLine("Server Error Message:" + exceptionHandlingScope.ErrorMessage);

                
                
            }
        }



    }



}
