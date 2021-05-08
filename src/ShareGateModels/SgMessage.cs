using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSTeamsHistory.ShareGateModels
{
    class SgMessage
    {
        [JsonProperty("subject")]
        public String Subject { get; set; }

        [JsonProperty("body")]
        public SgMessageBody Body { get; set; }

        [JsonProperty("attachments")]
        public List<SgAttachment> Attachments { get; set; }

        [JsonProperty("mentions")]
        public List<object> Mentions { get; set; }

        [JsonProperty("importance")]
        public String Importance { get; set; }
    }
}
