using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;



namespace Excel2DTDL
{
    public class DTDLModel
    {
        public Interface[] InterfaceArray { get; set; }
    }

    public class Interface
    {
        [JsonProperty(PropertyName = "@type")]
        public string type { get; set; }
        [JsonProperty(PropertyName = "@id")]
        public string id { get; set; }
        [JsonProperty(PropertyName = "@name")]
        public string name { get; set; }
#nullable enable
        public string? displayName { get; set; }
        public string? extends { get; set; }
        public string? description { get; set; }
#nullable disable
        [JsonProperty(PropertyName = "@context")]
        public string context { get; set; }
        public Content[] contents { get; set; }
    }

    public class Content
    {
        [JsonProperty(PropertyName = "@type")]
        public string[] type { get; set; }
        public string name { get; set; }
        public string schema { get; set; }
    }

    public class Property : Content
    {
#nullable enable
        public string? unit { get; set; }
        public bool? writable { get; set; }
#nullable disable
    }

    public class Telemetry : Content
    {
#nullable enable
        public string? unit { get; set; }
#nullable disable
    }

    public class Component : Content
    {
#nullable enable
        public string? displayName { get; set; }
        public string? description { get; set; }
#nullable disable
    }

    public class RelationShip : Content
    {
#nullable enable
        public string? displayName { get; set; }
#nullable disable
        public string target { get; set; }
    }

}
