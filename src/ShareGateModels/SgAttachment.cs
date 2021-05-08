using System;
using Newtonsoft.Json;

namespace MSTeamsHistory.ShareGateModels
{
    public class SgAttachment
    {
        [JsonProperty("id")]
        public String Id { get; set; }

        [JsonProperty("contentType")]
        public String ContentType { get; set; }

        [JsonProperty("content")]
        public String Content { get; set; }

        [JsonProperty("name")]
        public String Name { get; set; }

        [JsonProperty("exportedAttachmentContentUrl")]
        public String ExportedAttachmentContentUrl { get; set; }
    }
}
