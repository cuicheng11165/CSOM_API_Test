using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;

namespace CSOM_File_Add_Test
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, errors) => true;

            Write(AddFileWithBytes);
            Write(AddFileWithStream);
            Write(AddFileWithSaveBinaryDirect);
            Write(AddFileWithSaveBytes);
            Write(AddFileWithSaveStream);

            Write(AddLargeFileWithStream);
            Write(AddLargeFileWithSaveBinaryDirect);
        }


        public static void Write(Action testDeleagate)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            testDeleagate.Invoke();
            stopwatch.Stop();
            Console.WriteLine("TimeUsed :" + stopwatch.ElapsedMilliseconds);
        }

        private static void AddFileWithBytes()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);

            var web = context.Web;

            Folder folder = web.GetFolderByServerRelativeUrl("https://cnblogtest.sharepoint.com/Documents%20Test");

            var newAddedFile = folder.Files.Add(new FileCreationInformation()
            {
                Url = "AddFileWithBytes.txt",
                Overwrite = true,
                Content = Encoding.UTF8.GetBytes("TestDocumentContent")
            });

            context.Load(newAddedFile);
            context.ExecuteQuery();
        }

        private static void AddFileWithStream()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);

            var web = context.Web;

            Folder folder = web.GetFolderByServerRelativeUrl("https://cnblogtest.sharepoint.com/Documents%20Test");

            var newAddedFile = folder.Files.Add(new FileCreationInformation()
            {
                Url = "AddFileWithStream.txt",
                Overwrite = true,
                ContentStream = new MemoryStream(Encoding.UTF8.GetBytes("TestDocumentContent"))
            });

            context.Load(newAddedFile);
            context.ExecuteQuery();
        }

        private static void AddLargeFileWithStream()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);

            var web = context.Web;

            Folder folder = web.GetFolderByServerRelativeUrl("https://cnblogtest.sharepoint.com/Documents%20Test");

            using (FileStream fs = new FileStream("d:\\TestObject.rar", FileMode.Open))
            {

                var newAddedFile = folder.Files.Add(new FileCreationInformation()
                {
                    Url = "AddFileWithStreamLarge.rar",
                    Overwrite = true,
                    ContentStream = fs
                });

                context.Load(newAddedFile);
                context.ExecuteQuery();

            }
        }

        private static void AddFileWithSaveBinaryDirect()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);


            File.SaveBinaryDirect(context, "/Documents%20Test/AddFileWithSaveBinaryDirectLarge.rar", new MemoryStream(Encoding.UTF8.GetBytes("TestDocumentContent")), true);

            context.ExecuteQuery();
        }


        private static void AddLargeFileWithSaveBinaryDirect()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);

            using (FileStream fs = new FileStream("d:\\TestObject.rar", FileMode.Open))
            {
                File.SaveBinaryDirect(context, "/Documents%20Test/AddFileWithSaveBinaryDirectLarge.rar", fs, true);
            }
            context.ExecuteQuery();
        }

        private static void AddFileWithSaveBytes()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);

            var web = context.Web;

            var file = web.GetFileByServerRelativeUrl("/Documents%20Test/AddFileWithSaveBytes.txt");
            file.SaveBinary(new FileSaveBinaryInformation()
            {
                Content = Encoding.UTF8.GetBytes("TestDocumentContent")
            });


            context.ExecuteQuery();
        }

        private static void AddFileWithSaveStream()
        {
            ClientContext context = new ClientContext("https://cnblogtest.sharepoint.com");

            SecureString passworSecureString = new SecureString();

            "abc123!@#".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("test001@cnblogtest.onmicrosoft.com", passworSecureString);

            var web = context.Web;

            var file = web.GetFileByServerRelativeUrl("/Documents%20Test/AddFileWithSaveStream.txt");
            file.SaveBinary(new FileSaveBinaryInformation()
            {
                ContentStream = new MemoryStream(Encoding.UTF8.GetBytes("TestDocumentContent"))
            });

            context.ExecuteQuery();
        }
    }
}
