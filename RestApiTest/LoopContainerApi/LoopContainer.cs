using CSOM.Common;
class LoopContainer : RestApibase
{
    public void AddSPOContainerUserRole()
    {
        var adminUrl = EnvConfig.GetAdminCenterUrl();

        var token = EnvConfig.GetCsomToken();

        var url = $"{adminUrl}/_api/SPO.Tenant/AddSPOContainerUserRole";


        var body = new UserRoleInfo
        {
            ContainerId = "b!cbVG9fN3x0a4y-SRzz4ygBN1rEsDORVIoVoqpy-MK7iYCCQMZhJ6T5h-mZKqu459",
            loginName = "Eric@cloudgov.onmicrosoft.com",
            role = "owner"
        };

        SendPostRequestAsync(url, token, new StringContent(body.ToString())).GetAwaiter().GetResult();



    }


    public void RemoveSPOContainerUserRole()
    {
        var adminUrl = EnvConfig.GetAdminCenterUrl();

        var token = EnvConfig.GetCsomToken();

        var url = $"{adminUrl}/_api/SPO.Tenant/RemoveSPOContainerUserRole";


        var body = new UserRoleInfo
        {
            ContainerId = "b!cbVG9fN3x0a4y-SRzz4ygBN1rEsDORVIoVoqpy-MK7iYCCQMZhJ6T5h-mZKqu459",
            loginName = "Eric@cloudgov.onmicrosoft.com",
            role = "owner"
        };

        SendPostRequestAsync(url, token, new StringContent(body.ToString())).GetAwaiter().GetResult();



    }


}
