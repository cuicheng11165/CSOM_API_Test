#load "csomHelper.csx"

using System.Net;
using System;
using System.Security;
using System.Net.Http.Headers;

using OfficeDevPnP.Core;  
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.Identity.Client;

using Microsoft.Online.SharePoint.TenantAdministration;  
using Microsoft.Online.SharePoint.TenantManagement;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public class ContentTypePayload
{
    public string contentTypeId { get; set; }
}

public static async Task<HttpResponseMessage> 
Run(HttpRequestMessage req, TraceWriter log, ExecutionContext context)
{
        //TODO: 
        //Enable the Records365 event receivers...

    try { 
        log.Info("Executing function");

        string location = context.FunctionDirectory;
 
        dynamic data = await req.Content.ReadAsAsync<object>();
       
        //convert the Newtonsoft.Json.Linq.JObject to a dictionary     
        var Metadata = data?.Metadata.ToObject<Dictionary<string, object>>();

        //log the request details
        string RequestName = (Metadata.ContainsKey("RequestSummary") ? Metadata["RequestSummary"] : string.Empty);
        log.Info("Request Name: " + RequestName);
        long RequestId = (Metadata.ContainsKey("RequestId") ? Metadata["RequestId"] : 0);
        log.Info("Request ID: " + RequestId);
        string Requester = (Metadata.ContainsKey("Requester") ? Metadata["Requester"] : string.Empty);
        log.Info("Requester: " + Requester);

        log.Info("Got Metadata");
 
        string SiteTitle = (Metadata.ContainsKey("SiteTitle") ? Metadata["SiteTitle"] : string.Empty);
        log.Info("SiteTitle: " + SiteTitle);
        string SiteDescription = (Metadata.ContainsKey("SiteDescription") ? Metadata["SiteDescription"] : string.Empty);
        log.Info("SiteDescription: " + SiteDescription);
        
        string siteUrl = "";
        var UrlDetails = (Metadata.ContainsKey("Url") ? Metadata["Url"] : string.Empty);
        if (UrlDetails != null)
            siteUrl = UrlDetails.Value<string>("Prefix") + UrlDetails.Value<string>("ManagedPath") + UrlDetails.Value<string>("Url");

        log.Info(siteUrl);

        string custodian = (Metadata.ContainsKey("PrimaryContact") ? Metadata["PrimaryContact"] : string.Empty);
        custodian = custodian.TrimEnd(';');
        log.Info("Custodian: " + custodian);
        string champion = (Metadata.ContainsKey("SecondaryContact") ? Metadata["SecondaryContact"] : string.Empty); 
        champion = champion.TrimEnd(';');
        log.Info("Champion: " + champion);
        string consultant = (Metadata.ContainsKey("Consultant") ? Metadata["Consultant"] : string.Empty); 
        consultant = consultant.TrimEnd(';');
        log.Info("Consultant: " + consultant);

        string DepartmentArea = (Metadata.ContainsKey("Department Area") ? Metadata["Department Area"] : string.Empty); 
        DepartmentArea = DepartmentArea.TrimEnd(';');
        log.Info("Department Area: " + DepartmentArea);
    

        Uri itemUri = new Uri(siteUrl);
        var serverRelativeUrl = itemUri.AbsolutePath.ToString().TrimEnd('/');
        var pageRelativeUrl = serverRelativeUrl + "/SitePages/Home.aspx";

        using (var ctx = await csomHelper.GetClientContext(siteUrl, location))
        {
            //run this first, some bug with PnP not finding sharepoint dll
            var list = ctx.Web.GetListByTitle("Site Pages");

            var termgroupname = "Department Structure";
            //var termsetname = "Structure";

            var Structure = new Dictionary<string,string[]>();

            var RootSite = System.Environment.GetEnvironmentVariable("Adalink", EnvironmentVariableTarget.Process);
            using (var rootCtx = await csomHelper.GetClientContext(RootSite, location))
            { 
                //loading here as sometimes there was inconsistent behaviour in the next
                try {var tGroup = Microsoft.SharePoint.Client.TaxonomyExtensions.GetTermGroupByName(ctx.Site, termgroupname); }
                catch { log.Error("Error getting the term group"); }

// Removed the GetDepartmentStructure call as this creates too many errors. Alternative fix being implemented - DT 19-12-2022
//                Structure = csomHelper.GetDepartmentStructure(ctx, termgroupname, termsetname, DepartmentArea);
            }

            //agency not yet implemented - string agency = (Structure.ContainsKey("Agency") ? Structure["Agency"] : string.Empty);
            string group1 = (Structure.ContainsKey("Group") ? Structure["Group"][0] : string.Empty); 
            string group1Id = (Structure.ContainsKey("Group") ? Structure["Group"][1] : string.Empty); 
            string division = (Structure.ContainsKey("Division") ? Structure["Division"][0] : string.Empty); 
            string divisionId = (Structure.ContainsKey("Division") ? Structure["Division"][1] : string.Empty); 
            string branch = (Structure.ContainsKey("Branch") ? Structure["Branch"][0] : string.Empty); 
            string branchId = (Structure.ContainsKey("Branch") ? Structure["Branch"][1] : string.Empty);  
            string section = (Structure.ContainsKey("Section") ? Structure["Section"][0] : string.Empty); 
            string sectionId = (Structure.ContainsKey("Section") ? Structure["Section"][1] : string.Empty);

            //subsection not yet implemented - string subsection = (Structure.ContainsKey("SubSection") ? Structure["SubSection"] : string.Empty);
            
            ctx.Load(ctx.Web);
            ctx.ExecuteQuery();  

            log.Info("Set property bags");
            ctx.Web.SetPropertyBagValue("Custodian", custodian);
            ctx.Web.SetPropertyBagValue("Champion", champion);
            ctx.Web.SetPropertyBagValue("Consultant", consultant);
            ctx.Web.SetPropertyBagValue("Group", group1);
            ctx.Web.SetPropertyBagValue("Division", division);
            ctx.Web.SetPropertyBagValue("Branch", branch);
            ctx.Web.SetPropertyBagValue("Section", section);
            ctx.Web.SetPropertyBagValue("DepartmentArea", DepartmentArea);
            ctx.Web.SetPropertyBagValue("_auditlogreportstoragelocation", ctx.Web.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/MaintenanceLogs"); 

            log.Info("Deactivating Features");
            
            log.Info("getting started");
            ctx.Web.DeactivateFeature(new Guid("4aec7207-0d02-4f4f-aa07-b370199cd0c7"));
            log.Info("site notebook");
            ctx.Web.DeactivateFeature(new Guid("f151bb39-7c3b-414f-bb36-6bf18872052f"));


            log.Info("Activating Features");
            log.Info("Document ID");
            ctx.Site.ActivateFeature(new Guid("b50e3104-6812-424f-a011-cc90e6327318"),false,5);
            log.Info("Managed Metadata Navigation");
            ctx.Web.ActivateFeature(new Guid("7201d6a4-a5d3-49a1-8c19-19c4bac6e668"));
            log.Info("Site Feed");
            ctx.Web.ActivateFeature(new Guid("15a572c6-e545-4d32-897a-bab6f5846e18"));
            //No longer activate the Limited-access user permission lockdown mode feature 3-July-2018
            //log.Info("Limited-access user permission lockdown mode");
            //ctx.Site.ActivateFeature(new Guid("7c637b23-06c4-472d-9a9a-7c175762c5c4"));


            //enable auditing
            ctx.Load(ctx.Site);
            ctx.Load(ctx.Web);
            ctx.Load(ctx.Site.Audit);
            ctx.ExecuteQuery();
            ctx.Site.Audit.AuditFlags = AuditMaskType.All;
            ctx.Site.Audit.Update();

            //Line commented out by DELWP due to Error - Refer INC0018741
            //csomHelper.EnsureSiteAdministrators(ctx);

            string prefix = itemUri.AbsolutePath.ToString().TrimEnd('/').ToLower().Replace("/sites/","").Replace("ecm_", "").Replace("ecmtest_", "Test");
            ctx.Web.SetPropertyBagValue("docid_msft_hier_siteprefix", "DOCID" + prefix);
            log.Info("set doc id prefix to" + prefix);


            log.Info("Set Navigation");      
            ctx.Web.QuickLaunchEnabled = true;
            ctx.Web.TreeViewEnabled = false;
            ctx.Web.DeleteAllNavigationNodes(OfficeDevPnP.Core.Enums.NavigationType.QuickLaunch);
            ctx.Web.DeleteAllNavigationNodes(OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar);
            
            log.Info("Web Title:");
            log.Info(ctx.Web.Title);

            //Commented out on request by DELWP
            //var AdaLink = System.Environment.GetEnvironmentVariable("Adalink", EnvironmentVariableTarget.Process);
            //ctx.Web.AddNavigationNode("Ada", new Uri(AdaLink), string.Empty, OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar, false, true);
            NavigationExtensions.AddNavigationNode(ctx.Web, ctx.Web.Title, itemUri, string.Empty, OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar, false, false);
            var ECMSupportLink = System.Environment.GetEnvironmentVariable("ECMSupportlink", EnvironmentVariableTarget.Process);
            ctx.Web.AddNavigationNode("Support", new Uri(ECMSupportLink), string.Empty, OfficeDevPnP.Core.Enums.NavigationType.TopNavigationBar, false, true);
            
            //Quick launch home link
            NavigationExtensions.AddNavigationNode(ctx.Web, "Home", itemUri, string.Empty, OfficeDevPnP.Core.Enums.NavigationType.QuickLaunch, false, false);
            ctx.ExecuteQuery();

            log.Info("Delete Documents Library");
            if (ListExtensions.ListExists(ctx.Web, "Documents"))
                ListExtensions.GetListByTitle(ctx.Web, "Documents").DeleteObject();

            log.Info("remove how to use this library");
            List sitePages = ListExtensions.GetListByTitle(ctx.Web, "Site Pages");
            var file = FileFolderExtensions.GetFile(sitePages.RootFolder, "How To Use This Library.aspx");
            if (file != null) file.DeleteObject();

            log.Info("hide site assets from search");
            var listAssets = ListExtensions.GetListByTitle(ctx.Web, "Site Assets");
            listAssets.NoCrawl = true;
            listAssets.Update();
            
            log.Info("set the search settings");
            SearchExtensions.SetSiteCollectionSearchCenterUrl(ctx.Web, "");
            ctx.Web.SetPropertyBagValue("SRCH_SB_SET_SITE", "{\"Inherit\":true,\"ResultsPageAddress\":\"\",\"ShowNavigation\":true}");

            log.Info("execute");
            ctx.ExecuteQuery();

            log.Info("Set the search dropdown options");
            csomHelper.SetSearchDropdown(ctx);


            log.Info("delete the onenote notebook");
            csomHelper.RemoveSiteNotebook(ctx);

            log.Info("remove unused web parts");
            ctx.Web.DeleteWebPart(pageRelativeUrl, "Get started with your site");
            ctx.Web.DeleteWebPart(pageRelativeUrl, "Site feed");
            ctx.Web.DeleteWebPart(pageRelativeUrl, "Documents");

            log.Info("add new web parts");
            log.Info("PageUrl: " + pageRelativeUrl);
            csomHelper.EnsureContactWebPart(ctx.Web, pageRelativeUrl, "Site Custodian", custodian);
            csomHelper.EnsureContactWebPart(ctx.Web, pageRelativeUrl, "ECM Champion", champion);
            csomHelper.EnsureUsefulLinksWebPart(ctx.Web, pageRelativeUrl, "Useful Links");
            csomHelper.EnsureGAWebPart(ctx.Web, pageRelativeUrl, "SiteInfoCard");

            log.Info("set the regional settings");
            var tz = ctx.Web.RegionalSettings.TimeZones.GetById(76);
            ctx.Load(tz);
            ctx.Web.RegionalSettings.TimeZone = tz;
            ctx.Web.RegionalSettings.LocaleId = 3081;
            ctx.Web.RegionalSettings.FirstDayOfWeek = 1;
            ctx.Web.RegionalSettings.Time24 = true;
            ctx.Web.RequestAccessEmail = string.Empty;
            ctx.Web.Update();
            ctx.ExecuteQuery();

            //Disable the SharePoint Designer Function
            log.Info("Disable SharePoint Designer");
            ctx.Site.AllowDesigner = false;
            ctx.Site.AllowMasterPageEditing = false;
            ctx.Site.AllowRevertFromTemplate = false;
            ctx.Site.ShowUrlStructure = false;
            ctx.ExecuteQuery(); 
         

            //set column default values
            log.Info("set column default values");
 //           SetDefaultValue(ctx, "Branch", branchId , branch, log); 
 //           SetDefaultValue(ctx, "Division", divisionId , division, log); 
 //           SetDefaultValue(ctx, "Group1", group1Id , group1, log); 
 //           SetDefaultValue(ctx, "Section", sectionId, section, log); 


            //ensure that the permission levels exist
            log.Info("Creating ECM Permission levels");              
            csomHelper.EnsureRoleDefinition(ctx, "ECM Site Owner", "Owner level access to ECM Sites", csomHelper.GetOwnerPermissions(), log);
            csomHelper.EnsureRoleDefinition(ctx, "ECM Contributor", "Contribute level access to ECM Sites", csomHelper.GetMemberPermissions(), log);
            csomHelper.EnsureRoleDefinition(ctx, "ECM Viewer", "Read only access to ECM Sites", csomHelper.GetVisitorPermissions(), log);

            ctx.Load(ctx.Web.SiteGroups);
            ctx.Load(ctx.Web);
            ctx.ExecuteQuery();

            var ownersGroup = csomHelper.EnsureGroup(ctx, "Owners", "ECM Site Owner", log);
            var memberGroup = csomHelper.EnsureGroup(ctx, "Members", "ECM Contributor", log);
            var viewersGroup = csomHelper.EnsureGroup(ctx, "Visitors", "ECM Viewer", log);

            //Configure the owners for each permission group
            ownersGroup.Owner = ownersGroup;
            ownersGroup.Update();
            memberGroup.Owner = ownersGroup;
            memberGroup.Update();
            viewersGroup.Owner = ownersGroup;
            viewersGroup.Update();
            ctx.ExecuteQuery(); 

            ctx.Web.AssociateDefaultGroups(ownersGroup, memberGroup, viewersGroup);

            //guid for DELWP e8bdd6f7-fc18-4e48-a554-7f547927223b
            try {
                ctx.Web.AddUserToGroup(memberGroup, "c:0-.f|rolemanager|spo-grid-all-users/e8bdd6f7-fc18-4e48-a554-7f547927223b");
                ctx.Web.AddUserToGroup(ownersGroup, custodian.TrimEnd(';'));

                log.Info("add to the site directory");
                var ECMID = siteUrl.TrimEnd('/').Substring(siteUrl.Length - 1);
                await csomHelper.AddToSiteDirectory(siteUrl, SiteTitle, custodian, champion, ECMID, Structure, location);
            }
            catch (Exception ex)
            {
                string message = ex.Message + "\n" + ex.StackTrace;
                log.Error(message);
            }

            log.Info("Add content types from content type hub");
            var hostName = System.Environment.GetEnvironmentVariable("ResourceURL", EnvironmentVariableTarget.Process).Replace("https://","");
            string serverRelativePath = "";

            if (UrlDetails != null) {
                serverRelativePath = UrlDetails.Value<string>("ManagedPath") + UrlDetails.Value<string>("Url");
                var auth = await csomHelper.GetAuthentication(location);
                using (HttpClient client = new HttpClient()) 
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", auth.AccessToken);

                    var siteObj = await GetSiteId(client, hostName, serverRelativePath, log);
                    if (siteObj.Length > 0) {
                        var resultArray = siteObj.ToString().Split(',');
                        var siteId = resultArray[1].ToString();
                        log.Info($"Site ID: {siteId}");
                    
                        var result = await SyncContentTypes(client, siteId, log);
                    }
                    else {
                        log.Info("Failed to get site collection id from MS Graph.");
                    }
                }
            }
            else{
                log.Info("Url details not detected. Content Types were not added");
            }
        }

        return req.CreateResponse(HttpStatusCode.OK, "OK - Finished");
    }
    catch (Exception ex)
    {
        string message = ex.Message + "\n" + ex.StackTrace;
        log.Error(message);
        
        throw ex;
    }
}

private static async Task<string> GetSiteId(HttpClient client, string hostname, string serverRelativePath, TraceWriter log)
{
    try {
        using (HttpResponseMessage res = await client.GetAsync($"https://graph.microsoft.com/v1.0/sites/{hostname}:/{serverRelativePath}")) 
        {
            if (res.IsSuccessStatusCode) {
                var response = await res.Content.ReadAsStringAsync(); 
                dynamic siteObj = JObject.Parse(response);
                return siteObj.id.ToString();
            }
            else {
                var result = $"{(int)res.StatusCode} ({res.ReasonPhrase})";
                log.Info($"Failed: {result}");
                return "";
            }
        }
    }
    catch (Exception ex)
    {
        string message = ex.Message + "\n" + ex.StackTrace;
        log.Error(message);
        return "";
    }
}

private static async Task<bool> SyncContentTypes(HttpClient client, string siteId, TraceWriter log)
{
    try {
        var contentTypes = System.Environment.GetEnvironmentVariable("ContentTypesToSync", EnvironmentVariableTarget.Process);
        string[] contentTypesToSync = contentTypes.Split(';');

        using (HttpResponseMessage res = await client.GetAsync($"https://graph.microsoft.com/v1.0/sites/{siteId}/contentTypes/getCompatibleHubContentTypes")) 
        {
            if (res.IsSuccessStatusCode) {
                var str = await res.Content.ReadAsStringAsync(); 
                dynamic cts = JObject.Parse(str);

                foreach (var ct in cts.value) {
                    if (Array.IndexOf(contentTypesToSync, ct.group.ToString()) >= 0) {
                        var result = await CopyFromContentTypeHub(client, siteId, ct.id.ToString(), log);
                        if (result) {
                            log.Info($"Success. Name: {ct.name.ToString()}, Id: {ct.id.ToString()}");
                        }
                        else {
                            log.Info($"Failed. Name: {ct.name.ToString()}, Id: {ct.id.ToString()}");
                        }
                    }
                }
            }
            else {
                var result = $"{(int)res.StatusCode} ({res.ReasonPhrase})";
                log.Info($"Failed to retrieve list of content types: {result}"); 
            }
        }
        return true;
    }
    catch (Exception ex)
    {
        string message = ex.Message + "\n" + ex.StackTrace;
        log.Error(message);
        return false;
    }
}

private static async Task<bool> CopyFromContentTypeHub(HttpClient client, string siteId, string id, TraceWriter log)
{
    try {
        var content = new  ContentTypePayload() { contentTypeId = id };
        using (HttpResponseMessage res = await client.PostAsJsonAsync($"https://graph.microsoft.com/v1.0/sites/{siteId}/contentTypes/addCopyFromContentTypeHub", content))
        {
            if (res.IsSuccessStatusCode) {
                return true;
            }
            else {
                var result = $"{(int)res.StatusCode} ({res.ReasonPhrase})";
                log.Info($"Failed: {result}");
                return false;
            }
        }
    }
    catch (Exception ex)
    {
        string message = ex.Message + "\n" + ex.StackTrace;
        log.Error(message);
        return false;
    }
}

private static void SetDefaultValue(ClientContext ctx, string fieldInternalName, string termID, string termDisplayName, TraceWriter log) 
{
    try {
        if (string.IsNullOrEmpty(termDisplayName)) { return;}
        if (string.IsNullOrEmpty(termID)) { 
            throw new Exception(string.Format("Couldnt find id for column {0}", fieldInternalName));
        };

        string termDefaultValue = string.Format("-1;#{0}|{1}", termID, termDisplayName);

        log.Info(string.Format("Setting {0} to: {1}", fieldInternalName, termDefaultValue)); 

        ctx.Load(ctx.Web.Fields, flds => flds.Include(f => f.InternalName).Where(f => f.InternalName == fieldInternalName));
        ctx.ExecuteQuery();

        if (ctx.Web.Fields.Count > 0)
        {
            TaxonomyField field = ctx.CastTo<TaxonomyField>(ctx.Web.Fields.FirstOrDefault());

            var defaultValue = new TaxonomyFieldValue();
            defaultValue.WssId = -1;
            defaultValue.Label = termDisplayName;
            defaultValue.TermGuid = termID;

            //retrieve validated taxonomy field value
            var validatedValue = field.GetValidatedString(defaultValue);
            ctx.ExecuteQuery();

            //set default value for a taxonomy field
            field.DefaultValue = validatedValue.Value;
            field.Update();
            ctx.ExecuteQuery();
        } 
        else 
        { 
            throw new Exception("Could not find field: " + fieldInternalName);
        }
    }
    catch (Exception ex)
    {
        string message = ex.Message + "\n" + ex.StackTrace;
        log.Error(message);
    }
}  