/// <summary>
/// Code (mostly) generated using: https://app.quicktype.io/
/// </summary>
namespace Futurez.Xrm.DocumentTemplateSync
{
    using System;
    using System.Linq;
    using System.Collections.Generic;

    using System.Globalization;
    using System.IO;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public enum SelectedDocumentTypes
    {
        Both,
        Word,
        Excel
    }

    public partial class SyncSettings
    {
        [JsonProperty("connString_source", Required = Required.Default)]
        public string ConnStringSource { get; set; }

        public string ConnStringSourceUrl { get {
                var s = ConnStringSource.Split(';').Where(u=> u.ToLower().StartsWith("url=")).FirstOrDefault();
                return s;
            }
        }

        [JsonProperty("connString_target", Required = Required.Default)]
        public string ConnStringTarget { get; set; }
        public string ConnStringTargetUrl
        {
            get {
                var s = ConnStringTarget.Split(';').Where(u => u.ToLower().StartsWith("url=")).FirstOrDefault();
                return s;
            }
        }

        [JsonProperty("exportFolder", Required = Required.Default)]
        [JsonConverter(typeof(ParseStringConverter))]
        public string ExportFolder { get; set; }

        [JsonProperty("importFolder", Required = Required.Default)]
        [JsonConverter(typeof(ParseStringConverter))]
        public string ImportFolder { get; set; }

        [JsonProperty("exportSourceOnly", Required = Required.Always)]
        [JsonConverter(typeof(ParseStringConverter))]
        public bool ExportSourceOnly { get; set; }

        [JsonProperty("importTargetOnly", Required = Required.Always)]
        [JsonConverter(typeof(ParseStringConverter))]
        public bool ImportTargetOnly { get; set; }

        [JsonProperty("activeOnly", Required = Required.Always)]
        [JsonConverter(typeof(ParseStringConverter))]
        public bool ActiveOnly { get; set; }

        [JsonProperty("createOnly", Required = Required.Always)]
        [JsonConverter(typeof(ParseStringConverter))]
        public bool CreateOnly { get; set; }

        [JsonProperty("docTypes", Required = Required.Always)]
        [JsonConverter(typeof(ParseStringConverter))]
        public SelectedDocumentTypes DocumentTypes { get; set; }

        [JsonProperty("fileNameFilters", Required = Required.Default)]
        public List<string> FileNameFilters { get; set; }

        [JsonProperty("entityTypeFilters", Required = Required.Default)]
        public List<string> EntityTypeFilters { get; set; }

        [JsonProperty("loggingOnly", Required = Required.Default)]
        public bool LoggingOnly { get; set; }

        /// <summary>
        /// Look at the settings and make sure the depencies are setup
        /// </summary>
        public void DependencyCheck()
        {
            // connection string checks.
            // if export only then source is required
            if (ExportSourceOnly && ConnStringSource == null) {
                throw new ApplicationException("ExportSourceOnly specified but no ConnStringSource specified");
            }

            // if import only, then target and import folder is required
            if (ImportTargetOnly && ConnStringTarget == null) {
                throw new ApplicationException("ImportTargetOnly specified but no ConnStringTarget specified");
            }

            //// tell them about settings that might confuse...allow exit?
            //// target set but also export only
            //if (ImportTargetOnly && ConnStringSource != null) {
            //    DocumentSync.ReportProgress("NOTE: ImportTargetOnly specified but ConnStringSource specified. Is this correct?");
            //}

            //// target set but also export only 
            //if (ImportTargetOnly && ConnStringSource != null) {
            //    DocumentSync.ReportProgress("NOTE: ImportTargetOnly specified but ConnStringSource specified");
            //}


            //// both source and target set but folders are different
            //if (ConnStringTarget != null && ConnStringSource != null) {
            //}
        }
    }

    public partial class SyncSettings
    {
        public static SyncSettings FromFile(string[] args)
        {
            var json = "";
            var configFileName = "DocumentSyncConfig.json";
            var argsDict = new Dictionary<string, string>();

            // parse command line args... should only be one that names the config file.
            foreach (var arg in args) {
                var argDelimiterPosition = arg.IndexOf(":");
                var argName = arg.Substring(1, argDelimiterPosition - 1);
                var argValue = arg.Substring(argDelimiterPosition + 1);
                argsDict.Add(argName.ToLower(), argValue);

                DocumentSync.ReportProgress($"command line arg: Name: {argName.ToLower()}, Value: {argValue}");
            }

            if (!argsDict.ContainsKey("configfile")) {
                DocumentSync.ReportProgress($"Configruation file not passed on the command line. Looking for default in current directory: {configFileName}");
            }
            else {
                configFileName = argsDict["configfile"];
            }

            // load the config file
            if (File.Exists(configFileName)) {
                using (TextReader txtReader = new StreamReader(configFileName)) {
                    json = txtReader.ReadToEnd();
                }
            }
            else {
                throw new InvalidOperationException($"Unable to locate the configruation file {configFileName}.");
            }

            // now deserialize
            var settings = FromJson(json);

            // set the default import same as source.  they will export then import  
            if (settings.ConnStringTarget != null && settings.ImportFolder == null) {
                settings.ImportFolder = settings.ExportFolder;
            }

            return settings;
        }

        public static SyncSettings FromJson(string json) => JsonConvert.DeserializeObject<SyncSettings>(json, Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this SyncSettings self) => JsonConvert.SerializeObject(self, Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters = {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

    internal class ParseStringConverter : JsonConverter
    {
        public override bool CanConvert(Type t) => t == typeof(bool) || t == typeof(bool?) || t == typeof(SelectedDocumentTypes) || t == typeof(string);

        public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null) return null;
            var value = serializer.Deserialize<string>(reader);

            if (Boolean.TryParse(value, out bool b)) {
                return b;
            }

            if (Enum.TryParse<SelectedDocumentTypes>(value, out SelectedDocumentTypes s)) {
                return s;
            }

            // now just handle some string formatting.  cheesy, but saves a few lines of code
            // for now, we are just accounting for the datetime slugs
            if (t == typeof(string)) {
                return String.Format(value, DateTime.Now);
            }

            throw new Exception("Cannot parse incoming value");
        }

        public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer)
        {
            if (untypedValue == null) {
                serializer.Serialize(writer, null);
                return;
            }
            var value = (bool)untypedValue;
            var boolString = value ? "true" : "false";
            serializer.Serialize(writer, boolString);
            return;
        }

        public static readonly ParseStringConverter Singleton = new ParseStringConverter();
    }
}
