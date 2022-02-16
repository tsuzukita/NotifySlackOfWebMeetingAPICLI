using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

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
    }
}
