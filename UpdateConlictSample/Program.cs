using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SaveConflictTest
{
    class Program
    {
        static void Main(string[] args)
        {

            ClientContext context = new ClientContext("http://win-cpqm71buqvj:1000/sites/DC");

            var web = context.Web;
            var file1 = web.GetFileByServerRelativeUrl("/sites/DC/Test Document Id/ddd.docx");

            file1.CheckOut();

            file1.ListItemAllFields["FileLeafRef"] = "ddd.docx";
            file1.ListItemAllFields["Editor"] = 1;
            file1.ListItemAllFields["Author"] = 1;
            file1.ListItemAllFields["Modified"] = DateTime.UtcNow.AddDays(-1);
            file1.ListItemAllFields.Update();

            file1.CheckIn("", CheckinType.OverwriteCheckIn);

            file1.ListItemAllFields["FileLeafRef"] = "ddd.docx";
            file1.ListItemAllFields["Editor"] = 1;
            file1.ListItemAllFields["Author"] = 1;
            file1.ListItemAllFields["Modified"] = DateTime.Now.AddYears(1);

            file1.ListItemAllFields.Update();

            context.ExecuteQuery();//throws exception here 

            

            //          Microsoft.SharePoint.Client.ServerException was unhandled
            //HResult=-2146233088
            //Message=The file Test Document Id/ddd.docx has been modified by SHAREPOINT\system on 03 二月 2015 13:39:16 +0800.
            //Source=Microsoft.SharePoint.Client.Runtime
            //ServerErrorCode=-2130575305
            //ServerErrorTraceCorrelationId=ff01e69c-1de7-107b-db0e-21e25ad7e352
            //ServerErrorTypeName=Microsoft.SharePoint.SPException
            //ServerStackTrace=""
            //StackTrace:
            //     at Microsoft.SharePoint.Client.ClientRequest.ProcessResponseStream(Stream responseStream)
            //     at Microsoft.SharePoint.Client.ClientRequest.ProcessResponse()
            //     at Microsoft.SharePoint.Client.ClientRequest.ExecuteQueryToServer(ChunkStringBuilder sb)
            //     at Microsoft.SharePoint.Client.ClientRequest.ExecuteQuery()
            //     at Microsoft.SharePoint.Client.ClientRuntimeContext.ExecuteQuery()
            //     at Microsoft.SharePoint.Client.ClientContext.ExecuteQuery()
            //     at SaveConflictTest.Program.Main(String[] args) in c:\Users\ccui\Documents\GitHub\CSOM_API_Test\UpdateConlictSample\Program.cs:line 41
            //     at System.AppDomain._nExecuteAssembly(RuntimeAssembly assembly, String[] args)
            //     at System.AppDomain.ExecuteAssembly(String assemblyFile, Evidence assemblySecurity, String[] args)
            //     at Microsoft.VisualStudio.HostingProcess.HostProc.RunUsersAssembly()
            //     at System.Threading.ThreadHelper.ThreadStart_Context(Object state)
            //     at System.Threading.ExecutionContext.RunInternal(ExecutionContext executionContext, ContextCallback callback, Object state, Boolean preserveSyncCtx)
            //     at System.Threading.ExecutionContext.Run(ExecutionContext executionContext, ContextCallback callback, Object state, Boolean preserveSyncCtx)
            //     at System.Threading.ExecutionContext.Run(ExecutionContext executionContext, ContextCallback callback, Object state)
            //     at System.Threading.ThreadHelper.ThreadStart()
            //InnerException: 



        }

        public static void Update(Microsoft.SharePoint.Client.File file, string checkInComment, bool keepVersion)
        {
            MethodInfo updateMethod = typeof(ListItem).GetMethod("ValidateUpdateListItem", BindingFlags.Instance | BindingFlags.Public | BindingFlags.InvokeMethod);
            string fileLeafRef = file.ListItemAllFields.FieldValues.ContainsKey("FileLeafRef") ? file.ListItemAllFields["FileLeafRef"] as string : string.Empty;
            IList<ListItemFormUpdateValue> values = new List<ListItemFormUpdateValue>();
            values.Add(new ListItemFormUpdateValue() { FieldName = "FileLeafRef", FieldValue = fileLeafRef });
            if (updateMethod.GetParameters().Length == 3)
            {
                updateMethod.Invoke(file.ListItemAllFields, new object[] { values, keepVersion, checkInComment });
            }
            else
            {
                updateMethod.Invoke(file.ListItemAllFields, new object[] { values, keepVersion });
            }
        }
    }
}
