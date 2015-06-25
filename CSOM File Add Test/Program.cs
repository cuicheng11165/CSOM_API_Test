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

            AddFileWithContinueUpload();

            //Write(AddFileWithBytes);
            //Write(AddFileWithStream);
            //Write(AddFileWithSaveBinaryDirect);
            //Write(AddFileWithSaveBytes);
            //Write(AddFileWithSaveStream);

            //Write(AddLargeFileWithStream);
            //Write(AddLargeFileWithSaveBinaryDirect);
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

        private static void AddFileWithContinueUpload()
        {
            string fileName = "D:\\Reflector_8.3.0.93_XiaZaiBa.zip";



            ClientContext context = new ClientContext("https://simmon6.sharepoint.com/sites/Test1");

            SecureString passworSecureString = new SecureString();

            "1qaz2wsxE".ToCharArray().ToList().ForEach(passworSecureString.AppendChar);

            context.Credentials = new SharePointOnlineCredentials("cheng.cui@simmon6.onmicrosoft.com", passworSecureString);

            ClientResult<long> bytesUploaded = null;

            FileStream fs = null;
            File uploadFile = null;
            Guid uploadId = Guid.NewGuid();
            try
            {
                fs = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using (BinaryReader br = new BinaryReader(fs))
                {
                    byte[] buffer = new byte[1024 * 1024];
                    Byte[] lastBuffer = null;
                    long fileoffset = 0;
                    long totalBytesRead = 0;
                    int bytesRead;
                    bool first = true;
                    bool last = false;

                    // Read data from filesystem in blocks 
                    while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        totalBytesRead = totalBytesRead + bytesRead;

                        // We've reached the end of the file
                        if (totalBytesRead == fs.Length)
                        {
                            last = true;
                            // Copy to a new buffer that has the correct size
                            lastBuffer = new byte[bytesRead];
                            Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                        }

                        if (first)
                        {
                            using (MemoryStream contentStream = new MemoryStream())
                            {
                                // Add an empty file.
                                FileCreationInformation fileInfo = new FileCreationInformation();
                                fileInfo.ContentStream = contentStream;
                                fileInfo.Url = Path.GetFileName(fileName);
                                fileInfo.Overwrite = true;

                                Folder folder = context.Web.GetFolderByServerRelativeUrl("/sites/Test1/Test Library");


                                uploadFile = folder.Files.Add(fileInfo);

                                // Start upload by uploading the first slice. 
                                using (MemoryStream s = new MemoryStream(buffer))
                                {
                                    // Call the start upload method on the first slice
                                    bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                    context.ExecuteQuery();
                                    // fileoffset is the pointer where the next slice will be added
                                    fileoffset = bytesUploaded.Value;
                                    Console.WriteLine("fileoffset:" + fileoffset);
                                }

                                // we can only start the upload once
                                first = false;
                            }
                        }
                        else
                        {
                            // Get a reference to our file
                            uploadFile = context.Web.GetFileByServerRelativeUrl("/sites/Test1/Test Library" + System.IO.Path.AltDirectorySeparatorChar + Path.GetFileName(fileName));

                            if (last)
                            {
                                // Is this the last slice of data?
                                using (MemoryStream s = new MemoryStream(lastBuffer))
                                {
                                    // End sliced upload by calling FinishUpload
                                    uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                    context.ExecuteQuery();
                                    Console.WriteLine("fileoffset:" + fileoffset);
                                    // return the file object for the uploaded file
                                    break;
                                }
                            }
                            else
                            {
                                using (MemoryStream s = new MemoryStream(buffer))
                                {
                                    // Continue sliced upload
                                    bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                    context.ExecuteQuery();
                                    // update fileoffset for the next slice
                                    fileoffset = bytesUploaded.Value;
                                    Console.WriteLine("fileoffset:" + fileoffset);
                                }
                            }
                        }

                    } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                }
            }
            finally
            {
                //if (fs != null)
                //{
                //    fs.Dispose();
                //}
            }

            uploadFile = context.Web.GetFileByServerRelativeUrl("/sites/Test1/Test Library" + System.IO.Path.AltDirectorySeparatorChar + Path.GetFileName(fileName));

            ConditionalScope conditionScope = new ConditionalScope(context, () => uploadFile.Exists, true);
            using (conditionScope.StartScope())
            {
                using (conditionScope.StartIfTrue())
                {
                    context.Load(uploadFile);

                    context.Load(uploadFile.ListItemAllFields);
                }
            }
            context.ExecuteQuery();


        }

    }
}
