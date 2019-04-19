using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace RhumbixWPFMacro_KSE.ExcelData
{
    public partial class KseJson
    {
        [JsonProperty("job_number")]
        public string JobNumber { get; set; }

        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("store")]
        public Store Store { get; set; }

        [JsonProperty("schema")]
        public string Schema { get; set; }

        public string EmployeeId { get; set; }

        public int StartingRow { get; set; }

        public int Count { get; set; }

    }

    public partial class Store
    {
        [JsonProperty("Code", NullValueHandling = NullValueHandling.Ignore)]
        public Code Code { get; set; }

        [JsonProperty("Hours", NullValueHandling = NullValueHandling.Ignore)]
        public long? Hours { get; set; }

        [JsonProperty("Equipment", NullValueHandling = NullValueHandling.Ignore)]
        public Equipment Equipment { get; set; }

        [JsonProperty("Trade Change", NullValueHandling = NullValueHandling.Ignore)]
        public string TradeChange { get; set; }

        [JsonProperty("Cost Code Selector", NullValueHandling = NullValueHandling.Ignore)]
        public Code CostCodeSelector { get; set; }

        [JsonProperty("Hours as above trade on above cost code", NullValueHandling = NullValueHandling.Ignore)]
        public long? HoursAsAboveTradeOnAboveCostCode { get; set; }
    }

    public partial class Code
    {
        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("code")]
        public string CodeCode { get; set; }

        [JsonProperty("group")]
        public string Group { get; set; }

        [JsonProperty("label")]
        public string Label { get; set; }

        [JsonProperty("units")]
        public string Units { get; set; }

        [JsonProperty("project")]
        public Project Project { get; set; }

        [JsonProperty("is_active")]
        public bool IsActive { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }
    }

    public partial class Project
    {
        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("job_number")]
        public string JobNumber { get; set; }
    }

    public partial class Equipment
    {
        [JsonProperty("id")]
        public long Id { get; set; }

        [JsonProperty("label")]
        public string Label { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("project")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long Project { get; set; }

        [JsonProperty("category")]
        public string Category { get; set; }

        [JsonProperty("caltrans_id")]
        public string CaltransId { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("equipment_id")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long EquipmentId { get; set; }
    }

    public partial class KseJson
    {
        public static List<KseJson> FromJson(string json) => JsonConvert.DeserializeObject<List<KseJson>>(json, Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this List<KseJson> self) => JsonConvert.SerializeObject(self, Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

    internal class ParseStringConverter : JsonConverter
    {
        public override bool CanConvert(Type t) => t == typeof(long) || t == typeof(long?);

        public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null) return null;
            var value = serializer.Deserialize<string>(reader);
            long l;
            if (Int64.TryParse(value, out l))
            {
                return l;
            }
            throw new Exception("Cannot unmarshal type long");
        }

        public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer)
        {
            if (untypedValue == null)
            {
                serializer.Serialize(writer, null);
                return;
            }
            var value = (long)untypedValue;
            serializer.Serialize(writer, value.ToString());
            return;
        }

        public static readonly ParseStringConverter Singleton = new ParseStringConverter();
    }
}
