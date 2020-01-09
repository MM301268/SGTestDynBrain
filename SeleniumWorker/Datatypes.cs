using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SeleniumWorker
{
    public static class Datatypes
    {
        public struct WaitingMonitorStruct
        {
            public string Type { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string Status { get; set; }
            public string WatingTime { get; set; }
            public string Skill { get; set; }
            public string Queue { get; set; }
        }

    }

    public class RouteRequest
    {
        [JsonProperty(PropertyName = "external_id")]
        public string ExternalId { get; set; }
        [JsonProperty(PropertyName = "creation_timestamp")]
        public DateTimeOffset CreationTimestamp { get; set; }
        [JsonProperty(PropertyName = "tags")]
        public Dictionary<string, string> Tags { get; set; } = new Dictionary<string, string>();
    }

    public class RouteResponse
    {
        [JsonProperty(PropertyName = "contact_id")]
        public string ContactId { get; set; }
    }
}
