using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSTeamsHistory.Models.Graph
{
    class Metadata
    {
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [JsonProperty("@odata.mediaContentType")]
        public string ContentType { get; set; }

        [JsonProperty("@odata.mediaEtag")]
        public string etag { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("height")]
        public long Height { get; set; }

        [JsonProperty("width")]
        public long Width { get; set; }
    }
}
