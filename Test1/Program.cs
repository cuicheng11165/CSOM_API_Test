// See https://aka.ms/new-console-template for more information

using Microsoft.Azure.Services.AppAuthentication;

var ss =new AzureServiceTokenProvider();
var sss=ss.GetAccessTokenAsync("https://management.azure.com/").GetAwaiter().GetResult();


Console.ReadLine();