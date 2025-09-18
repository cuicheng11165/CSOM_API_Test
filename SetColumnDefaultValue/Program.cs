



 
using Microsoft.SharePoint.Client;
using System.Text;


ClientContext context = new ClientContext("https://bigapp.sharepoint.com/teams/Teams202504221153");

 

var list = context.Web.Lists.GetByTitle("lib6");
var column = list.Fields.GetByTitle("m1");

context.Load(column);
context.ExecuteQuery();

var termName = "C";
var termId = "23f5a117-458e-44fa-ac24-ff1fe1926054";
var defaultValueString = $"-1;#{termName}|{termId}";

column.DefaultValue = defaultValueString;

column.Update();

var newAddedFile = list.RootFolder.Files.Add(new FileCreationInformation()
{
    Url = "AddFileWithBytes.txt",
    Overwrite = true,
    Content = Encoding.UTF8.GetBytes("TestDocumentContent")
});

context.ExecuteQuery();





Console.ReadLine();