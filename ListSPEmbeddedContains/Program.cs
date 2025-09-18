using CSOM.Common;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;

var token = EnvConfig.GetToken();
var admin = EnvConfig.GetAdminCenterUrl();

var loopApplicationId = new Guid("a187e399-0c36-4b98-8f04-1edc167a0996");

ClientContext context = new ClientContext(admin);

context.ExecutingWebRequest += (object? sender, WebRequestEventArgs e) =>
{
    e.WebRequestExecutor.WebRequest.Headers[System.Net.HttpRequestHeader.Authorization] = token;
};

var tenant = new Tenant(context);
var containers = tenant.GetSPOContainersByApplicationId(loopApplicationId, false, "");
context.ExecuteQuery();

foreach (var containerProperty in containers.Value.ContainerCollection)
{
    var container = tenant.GetSPOContainerByContainerId(containerProperty.ContainerId);

    context.ExecuteQuery();

    Console.WriteLine($"AllowEditing: {containerProperty.AllowEditing}");

    Console.WriteLine($"AuthenticationContextName: {containerProperty.AuthenticationContextName}");
    Console.WriteLine($"BlockDownloadPolicy: {containerProperty.BlockDownloadPolicy}");
    Console.WriteLine($"ConditionalAccessPolicy: {containerProperty.ConditionalAccessPolicy}");
    Console.WriteLine($"ContainerApiUrl: {containerProperty.ContainerApiUrl}");
    Console.WriteLine($"ContainerId: {containerProperty.ContainerId}");
    Console.WriteLine($"ContainerName: {containerProperty.ContainerName}");
    Console.WriteLine($"ContainerSiteUrl: {containerProperty.ContainerSiteUrl}");
    Console.WriteLine($"ContainerTypeId: {containerProperty.ContainerTypeId}");
    Console.WriteLine($"CreatedBy: {containerProperty.CreatedBy}");
    Console.WriteLine($"CreatedOn: {containerProperty.CreatedOn}");
    Console.WriteLine($"Description: {containerProperty.Description}");
    Console.WriteLine($"ExcludeBlockDownloadPolicyContainerOwners: {containerProperty.ExcludeBlockDownloadPolicyContainerOwners}");
    Console.WriteLine($"LimitedAccessFileType: {containerProperty.LimitedAccessFileType}");
    Console.WriteLine($"Managers: {string.Join(",", containerProperty.Managers ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine($"Owners: {string.Join(",", containerProperty.Owners ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine($"OwnersCount: {containerProperty.OwnersCount}");
    Console.WriteLine($"OwningApplicationId: {containerProperty.OwningApplicationId}");
    Console.WriteLine($"OwningApplicationName: {containerProperty.OwningApplicationName}");
    Console.WriteLine($"Readers: {string.Join(",", containerProperty.Readers ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine($"ReadOnlyForBlockDownloadPolicy: {containerProperty.ReadOnlyForBlockDownloadPolicy}");
    Console.WriteLine($"ReadOnlyForUnmanagedDevices: {containerProperty.ReadOnlyForUnmanagedDevices}");
    Console.WriteLine($"SensitivityLabel: {containerProperty.SensitivityLabel}");
    Console.WriteLine($"SharingAllowedDomainList: {containerProperty.SharingAllowedDomainList}");
    Console.WriteLine($"SharingBlockedDomainList: {containerProperty.SharingBlockedDomainList}");
    Console.WriteLine($"SharingDomainRestrictionMode: {containerProperty.SharingDomainRestrictionMode}");
    Console.WriteLine($"Status: {containerProperty.Status}");
    Console.WriteLine($"StorageUsed: {containerProperty.StorageUsed}");
    Console.WriteLine($"Writers: {string.Join(",", containerProperty.Writers ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine("--------------------------------------------------");

}
Console.ReadLine();