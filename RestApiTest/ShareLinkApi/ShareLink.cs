using System.Text.Json;
using System.Text.Json.Serialization;
using CSOM.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RestApiTest.ShareLinkApi
{
    internal class ShareLink : RestApibase
    {
        public class ShareLinkRequest
        {
            [JsonPropertyName("request")]
            public RequestData? Request { get; set; }
        }

        public class RequestData
        {
            [JsonPropertyName("createLink")]
            public bool CreateLink { get; set; }

            [JsonPropertyName("settings")]
            public SettingsData? Settings { get; set; }
        }

        public class SettingsData
        {
            [JsonPropertyName("allowAnonymousAccess")]
            public bool AllowAnonymousAccess { get; set; }

            [JsonPropertyName("linkKind")]
            public int LinkKind { get; set; }

            [JsonPropertyName("expiration")]
            public object? Expiration { get; set; }

            [JsonPropertyName("role")]
            public int Role { get; set; }

            [JsonPropertyName("restrictShareMembership")]
            public bool RestrictShareMembership { get; set; }

            [JsonPropertyName("updatePassword")]
            public bool UpdatePassword { get; set; }

            [JsonPropertyName("password")]
            public string? Password { get; set; }

            [JsonPropertyName("scope")]
            public int Scope { get; set; }

            [JsonPropertyName("nav")]
            public string? Nav { get; set; }
        }

        public void CreateShareLink(string siteUrl, Guid listId, int itemId)
        {
            // Initialize SettingsData with explicit property values
            var settings = new SettingsData
            {
                AllowAnonymousAccess = false,
                LinkKind = 3,
                Expiration = null,
                Role = 2,
                RestrictShareMembership = false,
                UpdatePassword = false,
                Password = "",
                Scope = 1,
                Nav = ""
            };

            var requestData = new RequestData
            {
                CreateLink = true,
                Settings = settings
            };

            var shareLinkRequest = new ShareLinkRequest
            {
                Request = requestData
            };

            string json = JsonSerializer.Serialize(shareLinkRequest);


            // Encode dashes in listId as %2D for the URL
            string encodedListId = listId.ToString();
            string url = $"{siteUrl}/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1='{encodedListId}'&@a2='{itemId}'";
            string token = EnvConfig.GetCsomToken();
           this.SendPostRequestAsync(url, token, new StringContent(json)).Wait() ;
        }

    }
}
