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
        
            context.ExecuteQuery();        

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
