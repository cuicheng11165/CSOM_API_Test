using CSOM.Common;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;

var token = EnvConfig.GetCsomToken();
var admin = EnvConfig.GetAdminCenterUrl();

var loopApplicationId = new Guid("a187e399-0c36-4b98-8f04-1edc167a0996");

ClientContext context = new ClientContext(admin);

context.ExecutingWebRequest += (object? sender, WebRequestEventArgs e) =>
{
    e.WebRequestExecutor.WebRequest.Headers[System.Net.HttpRequestHeader.Authorization] = token;
};

var tenant = new Tenant(context);

var consumingContainerTypes=tenant.GetSPOContainerTypes(SPContainerTypeTenantType.ConsumingTenant);
context.ExecuteQuery();

var owningContainerTypes=tenant.GetSPOContainerTypes(SPContainerTypeTenantType.OwningTenant);
context.ExecuteQuery();


var containers = tenant.GetSPOContainersByApplicationId(loopApplicationId, false, "");
context.ExecuteQuery();

foreach (var containerProperty in containers.Value.ContainerCollection)
{
    var container = tenant.GetSPOContainerByContainerId(containerProperty.ContainerId);

    context.ExecuteQuery();

    Console.WriteLine($"AllowEditing: {containerProperty.AllowEditing}");

    Console.WriteLine($"AuthenticationContextName: {container.Value.AuthenticationContextName}");
    Console.WriteLine($"BlockDownloadPolicy: {container.Value.BlockDownloadPolicy}");
    Console.WriteLine($"ConditionalAccessPolicy: {container.Value.ConditionalAccessPolicy}");
    Console.WriteLine($"ContainerApiUrl: {container.Value.ContainerApiUrl}");
    Console.WriteLine($"ContainerId: {container.Value.ContainerId}");
    Console.WriteLine($"ContainerName: {container.Value.ContainerName}");
    Console.WriteLine($"ContainerSiteUrl: {container.Value.ContainerSiteUrl}");
    Console.WriteLine($"ContainerTypeId: {container.Value.ContainerTypeId}");
    Console.WriteLine($"CreatedBy: {container.Value.CreatedBy}");
    Console.WriteLine($"CreatedOn: {container.Value.CreatedOn}");
    Console.WriteLine($"Description: {container.Value.Description}");
    Console.WriteLine($"ExcludeBlockDownloadPolicyContainerOwners: {container.Value.ExcludeBlockDownloadPolicyContainerOwners}");
    Console.WriteLine($"LimitedAccessFileType: {container.Value.LimitedAccessFileType}");
    Console.WriteLine($"Managers: {string.Join(",", container.Value.Managers ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine($"Owners: {string.Join(",", container.Value.Owners ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine($"OwnersCount: {container.Value.OwnersCount}");
    Console.WriteLine($"OwningApplicationId: {container.Value.OwningApplicationId}");
    Console.WriteLine($"OwningApplicationName: {container.Value.OwningApplicationName}");
    Console.WriteLine($"Readers: {string.Join(",", container.Value.Readers ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine($"ReadOnlyForBlockDownloadPolicy: {container.Value.ReadOnlyForBlockDownloadPolicy}");
    Console.WriteLine($"ReadOnlyForUnmanagedDevices: {container.Value.ReadOnlyForUnmanagedDevices}");
    Console.WriteLine($"SensitivityLabel: {container.Value.SensitivityLabel}");
    Console.WriteLine($"SharingAllowedDomainList: {container.Value.SharingAllowedDomainList}");
    Console.WriteLine($"SharingBlockedDomainList: {container.Value.SharingBlockedDomainList}");
    Console.WriteLine($"SharingDomainRestrictionMode: {container.Value.SharingDomainRestrictionMode}");
    Console.WriteLine($"Status: {container.Value.Status}");
    Console.WriteLine($"StorageUsed: {container.Value.StorageUsed}");
    Console.WriteLine($"Writers: {string.Join(",", container.Value.Writers ?? new System.Collections.Generic.List<string>())}");
    Console.WriteLine("--------------------------------------------------");

}
Console.ReadLine();