using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.Identity.Client;
 
public static class csomHelper { 
 
    private static string ClientId = System.Environment.GetEnvironmentVariable("AADClientID", EnvironmentVariableTarget.Process);
    private static string Cert = System.Environment.GetEnvironmentVariable("CertFileName", EnvironmentVariableTarget.Process);
    private static string CertPassword = System.Environment.GetEnvironmentVariable("CertPassword", EnvironmentVariableTarget.Process);
    private static string Authority = System.Environment.GetEnvironmentVariable("AuthorityUrl", EnvironmentVariableTarget.Process);
    private static string Resource = System.Environment.GetEnvironmentVariable("ResourceURL", EnvironmentVariableTarget.Process);
    
    public async static Task<ClientContext> GetClientContext(string siteUrl, string location)
    {

        var authenticationContext = new AuthenticationContext(Authority, false);
 
        //var certPath = Path.Combine(Environment.GetEnvironmentVariable("HOME"), "site\\wwwroot\\ConfigureSite0-2\\", Cert);
        var certPath = Path.Combine(location, Cert);
        var cert = new X509Certificate2(System.IO.File.ReadAllBytes(certPath),
        CertPassword,
        X509KeyStorageFlags.Exportable |
        X509KeyStorageFlags.MachineKeySet |
        X509KeyStorageFlags.PersistKeySet);
 
        var authenticationResult = await authenticationContext.AcquireTokenAsync(Resource, new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate(ClientId, cert));
        var token = authenticationResult.AccessToken;
        
        var ctx = new OfficeDevPnP.Core.PnPClientContext(siteUrl);
        ctx.ExecutingWebRequest += (s, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authenticationResult.AccessToken;
        };
        
        return ctx;
    }

    public async static Task<Microsoft.Identity.Client.AuthenticationResult> GetAuthentication(string location)
    {
        var certPath = Path.Combine(location, Cert);
        var cert = new X509Certificate2(System.IO.File.ReadAllBytes(certPath),
        CertPassword,
        X509KeyStorageFlags.Exportable |
        X509KeyStorageFlags.MachineKeySet |
        X509KeyStorageFlags.PersistKeySet);
        var app = ConfidentialClientApplicationBuilder
            .Create(ClientId)
            .WithCertificate(cert)
            .WithAuthority(new Uri(Authority))
            .Build();
        var scopes = new[] { "https://graph.microsoft.com/.default" };
        var authenticationResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();
        return authenticationResult;
    }

    public static bool RemoveSiteNotebook(ClientContext ctx)
    {

        var listAssets = ListExtensions.GetListByTitle(ctx.Web, "Site Assets");

        CamlQuery query = CamlQuery.CreateAllItemsQuery(100, "FileLeafRef");
        query.ViewXml = query.ViewXml.Replace("<ViewFields>", "<Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='Text'>Team Site Notebook</Value></Eq></Where></Query><ViewFields>");
        var items = listAssets.GetItems(query);
        ctx.Load(items);
        ctx.ExecuteQuery();

        foreach (var item in items)
        {
            listAssets.GetItemById(item.Id).DeleteObject();
            ctx.ExecuteQuery();
        }
        return true;
    }

    public static void EnsureSiteAdministrators(ClientContext ctx)
    {
        string CompanyAdministrator = System.Environment.GetEnvironmentVariable("CompanyAdministrator", EnvironmentVariableTarget.Process);
        string SharePointServiceAdministrator = System.Environment.GetEnvironmentVariable("SharePointServiceAdministrator", EnvironmentVariableTarget.Process);
        string SiteAdminEmail = System.Environment.GetEnvironmentVariable("SiteAdminEmail", EnvironmentVariableTarget.Process);
        
        List<OfficeDevPnP.Core.Entities.UserEntity> Admins = new List<OfficeDevPnP.Core.Entities.UserEntity>();

        Admins.Add(new OfficeDevPnP.Core.Entities.UserEntity
        {
            LoginName = CompanyAdministrator,
            Title = "Company Administrator"
        });

        Admins.Add(new OfficeDevPnP.Core.Entities.UserEntity
        {
            LoginName = SharePointServiceAdministrator,
            Title = "SharePoint Service Administrator"
        });
        ctx.Web.AddAdministrators(Admins);

        ctx.Web.RemoveAdministrator(new OfficeDevPnP.Core.Entities.UserEntity { LoginName = "i:0#.f|membership|" + SiteAdminEmail });

    }

    public static Group EnsureGroup(ClientContext ctx, string ECMGroupType, string ECMPermissionLevel, TraceWriter log)
    {
        Group currentGroup = null;

        string groupName = string.Format("{0} {1}", ctx.Web.Title, ECMGroupType);
        var groupDescription = string.Format("Use this group to grant people {0} permissions to the SharePoint site: {1}", ECMPermissionLevel.Replace("ECM ", "").ToLower(), ctx.Web.Title);

        var currentGroups = ctx.Web.SiteGroups.Where(g => g.Title.EndsWith(ECMGroupType));
        if (currentGroups.Count() == 0)
            currentGroup = ctx.Web.AddGroup(groupName, groupDescription, false);
        else
            currentGroup = currentGroups.FirstOrDefault();

        currentGroup.Title = groupName;
        currentGroup.Description = groupDescription;
        currentGroup.Update();
        ctx.ExecuteQuery();

        ctx.Load(currentGroup);
        ctx.ExecuteQuery();

        log.Info("Group Name: " + groupName);
        ctx.Web.AddPermissionLevelToGroup(currentGroup.Title, ECMPermissionLevel, true);

        return currentGroup;

    }


    public static void EnsureRoleDefinition(ClientContext ctx, string PermissionLevelName, string PermissionLevelDescription, BasePermissions basePermissions, TraceWriter log)
    {
        try {
            var roleDefs = ctx.Web.RoleDefinitions;
            ctx.Load(roleDefs);
            ctx.ExecuteQuery();

            var roleDef = roleDefs.Where(rd => rd.Name == PermissionLevelName);

            if (roleDef.Count() == 0)
            {
                ctx.Web.RoleDefinitions.Add(new RoleDefinitionCreationInformation
                {
                    Name = PermissionLevelName,
                    Description = PermissionLevelDescription,
                    BasePermissions = basePermissions

                });

                ctx.ExecuteQuery();
            }  
        }
        catch (Exception ex)
        {
            string message = ex.Message + "\n" + ex.StackTrace;
            log.Error(message);
        }
    }

    public static BasePermissions GetOwnerPermissions()
    {
        var basePermissions = new BasePermissions();
        basePermissions.Set(PermissionKind.ManageLists);
        basePermissions.Set(PermissionKind.CancelCheckout);
        basePermissions.Set(PermissionKind.AddListItems);
        basePermissions.Set(PermissionKind.EditListItems);
        basePermissions.Set(PermissionKind.DeleteListItems);
        basePermissions.Set(PermissionKind.ViewListItems);
        basePermissions.Set(PermissionKind.ApproveItems);
        basePermissions.Set(PermissionKind.OpenItems);
        basePermissions.Set(PermissionKind.ViewVersions);
        basePermissions.Set(PermissionKind.CreateAlerts);
        basePermissions.Set(PermissionKind.ViewFormPages);
        basePermissions.Set(PermissionKind.ManagePermissions);
        basePermissions.Set(PermissionKind.ViewUsageData);
        basePermissions.Set(PermissionKind.BrowseDirectories);
        basePermissions.Set(PermissionKind.ViewPages);
        basePermissions.Set(PermissionKind.EnumeratePermissions);
        basePermissions.Set(PermissionKind.BrowseUserInfo);
        basePermissions.Set(PermissionKind.ManageAlerts);
        basePermissions.Set(PermissionKind.UseRemoteAPIs);
        basePermissions.Set(PermissionKind.UseClientIntegration);
        basePermissions.Set(PermissionKind.Open);
        basePermissions.Set(PermissionKind.ManagePersonalViews);

        return basePermissions;
    }

    public static BasePermissions GetMemberPermissions()
    {
        var basePermissions = new BasePermissions();

        basePermissions.Set(PermissionKind.AddListItems);
        basePermissions.Set(PermissionKind.EditListItems);
        basePermissions.Set(PermissionKind.DeleteListItems);
        basePermissions.Set(PermissionKind.ViewListItems);
        basePermissions.Set(PermissionKind.ApproveItems);
        basePermissions.Set(PermissionKind.OpenItems);
        basePermissions.Set(PermissionKind.ViewVersions);
        basePermissions.Set(PermissionKind.CreateAlerts);
        basePermissions.Set(PermissionKind.ViewFormPages);
        basePermissions.Set(PermissionKind.ManagePermissions);
        basePermissions.Set(PermissionKind.BrowseDirectories);
        basePermissions.Set(PermissionKind.ViewPages);
        basePermissions.Set(PermissionKind.EnumeratePermissions);
        basePermissions.Set(PermissionKind.BrowseUserInfo);
        basePermissions.Set(PermissionKind.ManageAlerts);
        basePermissions.Set(PermissionKind.UseRemoteAPIs);
        basePermissions.Set(PermissionKind.UseClientIntegration);
        basePermissions.Set(PermissionKind.Open);
        basePermissions.Set(PermissionKind.ManagePersonalViews);

        return basePermissions;

    }
    
    public static BasePermissions GetVisitorPermissions()
    {
        var basePermissions = new BasePermissions();

        basePermissions.Set(PermissionKind.ViewListItems);
        basePermissions.Set(PermissionKind.OpenItems);
        basePermissions.Set(PermissionKind.ViewVersions);
        basePermissions.Set(PermissionKind.CreateAlerts);
        basePermissions.Set(PermissionKind.ViewFormPages);

        basePermissions.Set(PermissionKind.ViewPages);
        basePermissions.Set(PermissionKind.BrowseUserInfo);
        basePermissions.Set(PermissionKind.UseRemoteAPIs);
        basePermissions.Set(PermissionKind.UseClientIntegration);
        basePermissions.Set(PermissionKind.Open);

        return basePermissions;
    }

    public static void SetSearchDropdown(ClientContext ctx)
    {
        string Adalink = System.Environment.GetEnvironmentVariable("Adalink", EnvironmentVariableTarget.Process);
        
        string SearchThisSiteURL = System.Environment.GetEnvironmentVariable("SearchThisSiteURL", EnvironmentVariableTarget.Process);
        SearchThisSiteURL = ctx.Web.Url.TrimEnd('/') + "/" + SearchThisSiteURL.TrimStart('/');
        
        string SearchEverythingURL = System.Environment.GetEnvironmentVariable("SearchEverythingURL", EnvironmentVariableTarget.Process);
        SearchEverythingURL = Adalink.TrimEnd('/') + "/" + SearchEverythingURL.TrimStart('/');
        
        string SearchPeopleURL = System.Environment.GetEnvironmentVariable("SearchPeopleURL", EnvironmentVariableTarget.Process);
        SearchPeopleURL = Adalink.TrimEnd('/') + "/" + SearchPeopleURL.TrimStart('/');
        
        //string SearchConversationsURL = System.Environment.GetEnvironmentVariable("SearchConversationsURL", EnvironmentVariableTarget.Process);
        //SearchConversationsURL = Adalink.TrimEnd('/') + "/" + SearchConversationsURL.TrimStart('/');



        //Remove the search dropdown options
        var searchNav = ctx.Web.LoadSearchNavigation();
        List<NavigationNode> nodes = new List<NavigationNode>();
        for (int i = searchNav.Count; i > 0; i--)
            searchNav[i - 1].DeleteObject();
        ctx.ExecuteQuery();

        //set the search dropdown options
        NavigationExtensions.AddNavigationNode(ctx.Web, "This Site", new Uri(SearchThisSiteURL), string.Empty, OfficeDevPnP.Core.Enums.NavigationType.SearchNav, false, false);
        NavigationExtensions.AddNavigationNode(ctx.Web, "Everything", new Uri(SearchEverythingURL), string.Empty, OfficeDevPnP.Core.Enums.NavigationType.SearchNav, false, true);
        NavigationExtensions.AddNavigationNode(ctx.Web, "People", new Uri(SearchPeopleURL), string.Empty, OfficeDevPnP.Core.Enums.NavigationType.SearchNav, false, true);
        //NavigationExtensions.AddNavigationNode(ctx.Web, "Conversations", new Uri(SearchConversationsURL), string.Empty, OfficeDevPnP.Core.Enums.NavigationType.SearchNav, false, true);



    }


    public async static Task<bool> AddToSiteDirectory(string SiteUrl, string SiteName, string Custodian, string Champion, string ECMID, Dictionary<string, string[]>Structure, string location)
    {
        var SiteDirectoryURL = System.Environment.GetEnvironmentVariable("SiteDirectoryWebURL", EnvironmentVariableTarget.Process);
        var SiteDirectoryListName = System.Environment.GetEnvironmentVariable("SiteDirectoryListName", EnvironmentVariableTarget.Process);
        var SiteDirectoryDefaultStatus = System.Environment.GetEnvironmentVariable("SiteDirectoryStatus", EnvironmentVariableTarget.Process);

        using (var ctx = await csomHelper.GetClientContext(SiteDirectoryURL, location))
        {
            //get the username of the Custodiam, this is stored in the field as text
            string custodianName = string.Empty;
            var CustodianUser = ctx.Web.EnsureUser(Custodian.TrimEnd(';'));
            try { ctx.Load(CustodianUser); ctx.ExecuteQuery(); custodianName = CustodianUser.Title; } catch {  }

            //get the username of the Champion, this is stored in the field as text
            string championName = string.Empty;
            var ChampionUser = ctx.Web.EnsureUser(Champion.TrimEnd(';'));
            try { ctx.Load(ChampionUser); ctx.ExecuteQuery(); championName = ChampionUser.Title; } catch {  }

            var SiteDirectory = ctx.Web.GetListByTitle(SiteDirectoryListName);


            var items = GetCSOMListItemByField(SiteDirectory, "ECM_x0020_ID", ECMID);

/*
            //check and see if the item already exists
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRefName='ECM_x0020_ID'/><Eq><Value Type='Number'>"+ ECMID +"</Value></Eq></Where></Query></View>";    
            var itemColl = SiteDirectory.GetItems(camlQuery);
            ctx.Load(itemColl);
            ctx.ExecuteQuery();

            var count = itemColl.Count;
            */

            ListItem item;

            if (items.Count == 0){
                ListItemCreationInformation li = new ListItemCreationInformation();
                item = SiteDirectory.AddItem(li);
            } else item = items.FirstOrDefault();




           item["Title"] = (Structure.ContainsKey("Group") ? Structure["Group"][0] : string.Empty);
            item["ECM_x0020_ID"] = ECMID;

            var SiteCollectionURL = new FieldUrlValue();
            SiteCollectionURL.Url = SiteUrl;
            SiteCollectionURL.Description = SiteName;
            item["Site_x0020_Collection_x0020_Link"] = SiteCollectionURL;

            //this is the title
            //item["Group"] = Group;

           item["Division"] = (Structure.ContainsKey("Division") ? Structure["Division"][0] : string.Empty);
           item["Branch"] = (Structure.ContainsKey("Branch") ? Structure["Branch"][0] : string.Empty);
           item["Section"] = (Structure.ContainsKey("Section") ? Structure["Section"][0] : string.Empty);
            //not yet implemented
            //item["SubSection"] =

            item["ECM_x0020_Champion"] = championName;
            item["Site_x0020_Custodian"] = custodianName;

          //  var ChampionUser = ctx.Web.EnsureUser(Champion.TrimEnd(';'));
          //  try { ctx.ExecuteQuery(); item["ECM_x0020_Champion"] = ChampionUser.Title; } catch {  }
         
            item["site_status"] = SiteDirectoryDefaultStatus;

            item.Update();

            ctx.ExecuteQuery();

        }

        return true;
    }

    private static ListItemCollection GetCSOMListItemByField(List list, string FieldName, string FieldValue){

        var ctx = list.Context;

        var query = CamlQuery.CreateAllItemsQuery(1000, "Title");

        ctx.Load(list);
        ctx.Load(list.Fields);
        ctx.ExecuteQuery();

        var field = list.Fields.Where(f => f.InternalName == FieldName).FirstOrDefault();

        var QueryXML = string.Format("<Query><Where><Eq><FieldRef Name='{0}' /><Value Type='{1}'>{2}</Value></Eq></Where></Query><ViewFields>",
            FieldName, field.FieldTypeKind, FieldValue);



        //        var fieldsXML = "";
        //list.Fields | % { $fieldsXML += [string]::Format($XMLtemplate, $_.InternalName) }

        //query.ViewXml = query.ViewXml.Replace("<FieldRef Name=""Title"" />", $fieldsXML)
        query.ViewXml = query.ViewXml.Replace("<ViewFields>", QueryXML);


        var listItems = list.GetItems(query);
        ctx.Load(listItems);

        ctx.ExecuteQuery();

        return listItems;


    }

    public static Dictionary<string, string[]> GetDepartmentStructure(ClientContext ctx, string termGroupName, string termSetName, string termName) 
    { 

        var tGroup = ctx.Site.GetTermGroupByName(termGroupName);
        var Set = tGroup.TermSets.GetByName(termSetName);
        ctx.Load(Set);
        ctx.ExecuteQuery();

        var lmi = new LabelMatchInformation(ctx);
        lmi.TermLabel = termName;
        lmi.TrimUnavailable = true;

        var terms = Set.GetTerms(lmi);
        ctx.Load(terms);
        ctx.ExecuteQuery();

        var Structure = new Dictionary<string, string[]>();

        if (terms.Count > 0)
        {
            var term = terms[0];
            var termpath = term.PathOfTerm;
            var TermHireachy = termpath.Split(';');
            var NumberOfTerms = TermHireachy.Length;

            Term currentTerm;

            for (int i = NumberOfTerms-1; i > -1; i--)
            {

                var termValue = new string[2];
                termValue[0] = TermHireachy[i];
                //termValue[1] = currentTerm.Id.ToString();


                switch (i)
                {
                    case 0:
                        //termValue[1] = GetTermId(ctx, "ECM", "Agency", termValue),
                        Structure.Add("Agency", termValue);
                        break;
                    case 1:
                        termValue[1] = GetTermId(ctx, "ECM", "Group", termValue[0]);
                        //sometimes there is an issue getting the term. Try again
                        if (string.IsNullOrEmpty(termValue[1])) { System.Threading.Thread.Sleep(10000); termValue[1] = GetTermId(ctx, "ECM", "Group", termValue[0]);}
                        Structure.Add("Group", termValue);
                        break;
                    case 2:
                        termValue[1] = GetTermId(ctx, "ECM", "Division", termValue[0]);
                        //sometimes there is an issue getting the term. Try again
                        if (string.IsNullOrEmpty(termValue[1])) { System.Threading.Thread.Sleep(10000); termValue[1] = GetTermId(ctx, "ECM", "Division", termValue[0]);}
                        Structure.Add("Division", termValue);
                        break;
                    case 3:
                        termValue[1] = GetTermId(ctx, "ECM", "Branch", termValue[0]);
                        //sometimes there is an issue getting the term. Try again
                        if (string.IsNullOrEmpty(termValue[1])) { System.Threading.Thread.Sleep(10000); termValue[1] = GetTermId(ctx, "ECM", "Branch", termValue[0]);}
                        Structure.Add("Branch", termValue);
                        break;
                    case 4:
                        termValue[1] = GetTermId(ctx, "ECM", "Section", termValue[0]);
                        //sometimes there is an issue getting the term. Try again
                        if (string.IsNullOrEmpty(termValue[1])) { System.Threading.Thread.Sleep(10000); termValue[1] = GetTermId(ctx, "ECM", "Section", termValue[0]);}
                        Structure.Add("Section", termValue);
                        break;
                    case 5:
                        Structure.Add("SubSection", termValue);
                        break;
                }

                currentTerm = term.Parent;
                ctx.Load(currentTerm);
                ctx.ExecuteQuery();

            }
        }

        return Structure;



    }

    private static string GetTermId(ClientContext ctx, string termGroupName, string termSetName, string termName)
    {
        var tGroup = ctx.Site.GetTermGroupByName(termGroupName);
        var Set = tGroup.TermSets.GetByName(termSetName);
        ctx.Load(Set);
        ctx.ExecuteQuery();

        var lmi = new LabelMatchInformation(ctx);
        lmi.TermLabel = termName;
        lmi.TrimUnavailable = true;

        var terms = Set.GetTerms(lmi);
        ctx.Load(terms);
        ctx.ExecuteQuery();

        if (terms.Count == 0) throw new Exception(string.Format("Could not find a term with the name {0} in the term set {1} in the group {2}", termName, termSetName, termGroupName));

        var term = terms[0];

        ctx.Load(term);
        ctx.ExecuteQuery();

        return term.Id.ToString();

    }
   

    public static bool EnsureContactWebPart(Web web, string pageRelativeUrl, string labelText, string userEmail)
    {

        var WebParts = PageExtensions.GetWebParts(web, pageRelativeUrl);

        var ContactParts = WebParts.Where(wp => wp.WebPart.Properties.FieldValues.ContainsKey("LabelText"));
        var champParts = ContactParts.Where(wpd => (string)wpd.WebPart.Properties.FieldValues["LabelText"] == labelText);


        foreach (var champPart in champParts)
        {
            PageExtensions.DeleteWebPart(web, pageRelativeUrl, champPart.WebPart.Title);
        }

        web.EnsureUser(userEmail);
        try { web.Context.ExecuteQuery(); } catch { return false; }

        OfficeDevPnP.Core.Entities.WebPartEntity webPart = new OfficeDevPnP.Core.Entities.WebPartEntity();
        //webPart.WebPartTitle = "Custodian";
        webPart.WebPartXml = string.Format(@"<?xml version='1.0' encoding='utf-8'?>
                                <WebPart xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns='http://schemas.microsoft.com/WebPart/v2'>
                                    <Title>{0}</Title>
                                    <FrameType>None</FrameType>
                                    <Description>Displays details about a contact for this page or site.</Description>
                                    <IsIncluded>true</IsIncluded>
                                    <ZoneID>wpz</ZoneID>
                                    <PartOrder>0</PartOrder>
                                    <FrameState>Normal</FrameState>
                                    <Height />
                                    <Width />
                                    <AllowRemove>true</AllowRemove>
                                    <AllowZoneChange>true</AllowZoneChange>
                                    <AllowMinimize>true</AllowMinimize>
                                    <AllowConnect>true</AllowConnect>
                                    <AllowEdit>true</AllowEdit>
                                    <AllowHide>true</AllowHide>
                                    <IsVisible>true</IsVisible>
                                    <DetailLink />
                                    <HelpLink />
                                    <HelpMode>Modeless</HelpMode>
                                    <Dir>Default</Dir>
                                    <PartImageSmall />
                                    <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
                                    <PartImageLarge>/_layouts/images/wp_pers.gif</PartImageLarge>
                                    <IsIncludedFilter />
                                    <Assembly>Microsoft.SharePoint.Portal, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
                                    <TypeName>Microsoft.SharePoint.Portal.WebControls.ContactFieldControl</TypeName>
                                    <CacheTimeout xmlns='urn:schemas-microsoft-com:sharepoint:CacheableWebPart'>600</CacheTimeout>
                                    <ContactLoginName xmlns='urn:schemas-microsoft-com:contactfieldcontrol'>i:0#.f|membership|{1}</ContactLoginName>
                                    <LabelText xmlns='urn:schemas-microsoft-com:contactfieldcontrol'>{0}</LabelText>
                                </WebPart>", labelText, userEmail);
        PageExtensions.AddWebPartToWikiPage(web, pageRelativeUrl, webPart, 2, 2, true);
        
        return true;
    }

    public static bool EnsureGAWebPart(Web web, string pageRelativeUrl, string labelText)
    {

        var WebParts = PageExtensions.GetWebParts(web, pageRelativeUrl);

        //var champParts = WebParts.Where(wp => wp.WebPart.Title == "Site Chamption");
        var ContactParts = WebParts.Where(wp => wp.WebPart.Properties.FieldValues.ContainsKey("WebPartName"));
        var champParts = ContactParts.Where(wpd => (string)wpd.WebPart.Properties.FieldValues["WebPartName"] == labelText);


        foreach (var champPart in champParts)
            PageExtensions.DeleteWebPart(web, pageRelativeUrl, champPart.WebPart.Title);


        OfficeDevPnP.Core.Entities.WebPartEntity webPart = new OfficeDevPnP.Core.Entities.WebPartEntity();
        webPart.WebPartXml = @"<webParts>
                                <webPart xmlns='http://schemas.microsoft.com/WebPart/v3'>
                                <metaData>
                                    <type name='Microsoft.SharePoint.WebPartPages.ClientWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />
                                    <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
                                </metaData>
                                <data>
                                    <properties>
                                    <property name='TitleIconImageUrl' type='string' />
                                    <property name='Direction' type='direction'>NotSet</property>
                                    <property name='ExportMode' type='exportmode'>All</property>
                                    <property name='HelpUrl' type='string' />
                                    <property name='Hidden' type='bool'>False</property>
                                    <property name='Description' type='string'>Displays key information about the site such as ownership, classification, policy, and available service requests.</property>
                                    <property name='FeatureId' type='System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'>bb96d1bc-94dc-4fec-913a-2f2a8e995410</property>
                                    <property name='CatalogIconImageUrl' type='string' />
                                    <property name='Title' type='string'>Site Information Card</property>
                                    <property name='AllowHide' type='bool'>True</property>
                                    <property name='ProductWebId' type='System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'>7d6ac279-b97d-4a5d-b1fd-55010756e57a</property>
                                    <property name='AllowZoneChange' type='bool'>True</property>
                                    <property name='TitleUrl' type='string' />
                                    <property name='ChromeType' type='chrometype'>TitleOnly</property>
                                    <property name='AllowConnect' type='bool'>True</property>
                                    <property name='Width' type='unit'>300px</property>
                                    <property name='Height' type='unit'>350px</property>
                                    <property name='WebPartName' type='string'>SiteInfoCard</property>
                                    <property name='HelpMode' type='helpmode'>Navigate</property>
                                    <property name='AllowEdit' type='bool'>True</property>
                                    <property name='AllowMinimize' type='bool'>True</property>
                                    <property name='ProductId' type='System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'>bb96d1bc-94dc-4fec-913a-2f2a8e99540f</property>
                                    <property name='AllowClose' type='bool'>True</property>
                                    <property name='ChromeState' type='chromestate'>Normal</property>
                                    </properties>
                                </data>
                                </webPart>
                            </webParts>";

        PageExtensions.AddWebPartToWikiPage(web, pageRelativeUrl, webPart, 2, 2, true);

        return true;
    }

    public static bool EnsureUsefulLinksWebPart(Web web, string pageRelativeUrl, string labelText)
    {

        var WebParts = PageExtensions.GetWebParts(web, pageRelativeUrl);

        //var champParts = WebParts.Where(wp => wp.WebPart.Title == "Site Chamption");
        var ContactParts = WebParts.Where(wp => wp.WebPart.Properties.FieldValues.ContainsKey("Title"));
        var champParts = ContactParts.Where(wpd => (string)wpd.WebPart.Properties.FieldValues["Title"] == labelText); 


        foreach (var champPart in champParts)
            PageExtensions.DeleteWebPart(web, pageRelativeUrl, champPart.WebPart.Title);


        OfficeDevPnP.Core.Entities.WebPartEntity webPart = new OfficeDevPnP.Core.Entities.WebPartEntity();
        webPart.WebPartXml = @"<?xml version='1.0' encoding='utf-8'?>
                                    <WebPart xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns='http://schemas.microsoft.com/WebPart/v2'>
                                        <Title>Useful Links</Title>
                                        <FrameType>Default</FrameType>
                                        <Description>Allows authors to enter rich text content.</Description>
                                        <IsIncluded>true</IsIncluded>
                                        <ZoneID>wpz</ZoneID>
                                        <PartOrder>0</PartOrder>
                                        <FrameState>Normal</FrameState>
                                        <Height />
                                        <Width />
                                        <AllowRemove>true</AllowRemove>
                                        <AllowZoneChange>true</AllowZoneChange>
                                        <AllowMinimize>true</AllowMinimize>
                                        <AllowConnect>true</AllowConnect>
                                        <AllowEdit>true</AllowEdit>
                                        <AllowHide>true</AllowHide>
                                        <IsVisible>true</IsVisible>
                                        <DetailLink />
                                        <HelpLink />
                                        <HelpMode>Modeless</HelpMode>
                                        <Dir>Default</Dir>
                                        <PartImageSmall />
                                        <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
                                        <PartImageLarge>/_layouts/15/images/mscontl.gif</PartImageLarge>
                                        <IsIncludedFilter />
                                        <Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
                                        <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
                                        <ContentLink xmlns='http://schemas.microsoft.com/WebPart/v2/ContentEditor' />
                                        <Content xmlns='http://schemas.microsoft.com/WebPart/v2/ContentEditor' />
                                        <PartStorage xmlns='http://schemas.microsoft.com/WebPart/v2/ContentEditor' />
                                    </WebPart>";

        PageExtensions.AddWebPartToWikiPage(web, pageRelativeUrl, webPart, 2, 2, true);

        return true;
    }
}