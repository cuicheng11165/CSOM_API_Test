using System.Text.Json;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net.Http;
using RestApiTest.ShareLinkApi;
using CSOM.Common;

public class UserRoleInfo
{
    public string? ContainerId { get; set; }
    public string? loginName { get; set; }
    public string? role { get; set; }

    public override string ToString()
    {
        return JsonSerializer.Serialize(this);
    }
}

class RestApi
{


    static void Main()
    {
        var siteUrl = EnvConfig.GetSiteUrl("/contentstorage/x8FNO-xtskuCRX2_fMTHLT17vJaIE59ArxPpSSZt3Zw");
        new ShareLink().CreateShareLink(siteUrl, new Guid("0c240898-1266-4f7a-987e-9992aabb8e7d"), 5);

    }


}
