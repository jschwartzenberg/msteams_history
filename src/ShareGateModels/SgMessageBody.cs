using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSTeamsHistory.ShareGateModels
{
    class SgMessageBody
    {
        [JsonProperty("contentType")]
        public String ContentType { get; set; }

        [JsonProperty("content")]
        public String Content { get; set; }
    }
}
