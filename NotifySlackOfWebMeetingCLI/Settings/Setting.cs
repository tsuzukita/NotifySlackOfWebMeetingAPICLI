using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Web;
using Newtonsoft.Json;
using NotifySlackOfWebMeetingCLI.Authentication;
namespace NotifySlackOfWebMeetingCLI.Settings
{
    /// <summary>
    /// 設定
    /// </summary>
    public class Setting
    {
        /// <summary>
        /// SlackチェンネルのID
        /// </summary>
        [JsonProperty("slackChannelId")]
        public string SlackChannelId { get; set; }
        /// <summary>
        /// チャンネル名
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; }
        /// <summary>
        /// 登録者
        /// </summary>
        [JsonProperty("registeredBy")]
        public string RegisteredBy { get; set; }
        /// <summary>
        /// エンドポイントURL
        /// </summary>
        [JsonProperty("endpointUrl")]
        public string EndpointUrl { get; set; }
        /// <summary>
        /// instance of Azure AD, for example public Azure or a Sovereign cloud (Azure China, Germany, US government, etc ...)
        /// </summary>
        public string Instance { get; set; } = "https://login.microsoftonline.com/{0}";
        /// <summary>
        /// Graph API endpoint, could be public Azure (default) or a Sovereign cloud (US government, etc ...)
        /// </summary>
        public string ApiUrl { get; set; } = "https://graph.microsoft.com/";
        /// <summary>
        /// The Tenant is:
        /// - either the tenant ID of the Azure AD tenant in which this application is registered (a guid)
        /// or a domain name associated with the tenant
        /// - or 'organizations' (for a multi-tenant application)
        /// </summary>
        public string Tenant { get; set; }
        /// <summary>
        /// Guid used by the application to uniquely identify itself to Azure AD
        /// </summary>
        public string ClientId { get; set; }
        /// <summary>
        /// URL of the authority
        /// </summary>
        public string Authority
        {
            get
            {
                return String.Format(CultureInfo.InvariantCulture, Instance, Tenant);
            }
        }
        /// <summary>
        /// Client secret (application password)
        /// </summary>
        /// <remarks>Daemon applications can authenticate with AAD through two mechanisms: ClientSecret
        /// (which is a kind of application password: this property)
        /// or a certificate previously shared with AzureAD during the application registration
        /// (and identified by the Certificate property belows)
        /// <remarks>
        public string ClientSecret { get; set; }
        /// <summary>
        /// The description of the certificate to be used to authenticate your application.
        /// </summary>
        /// <remarks>Daemon applications can authenticate with AAD through two mechanisms: ClientSecret
        /// (which is a kind of application password: the property above)
        /// or a certificate previously shared with AzureAD during the application registration
        /// (and identified by this CertificateDescription)
        /// <remarks>
        public CertificateDescription Certificate { get; set; }
    }
}