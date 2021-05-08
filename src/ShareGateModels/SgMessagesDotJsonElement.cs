using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSTeamsHistory.ShareGateModels
{
    class SgMessagesDotJsonElement
    {
        [JsonProperty("message")]
        public String Message { get; set; }

        [JsonProperty("replies")]
        public List<string> Replies { get; set; }
    }
}
