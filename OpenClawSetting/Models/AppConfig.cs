using System.Text.Json.Serialization;


namespace OpenClawSetting.Models
{
    public class AuthProfileConfig
    {
        [JsonPropertyName("version")]
        public int Version { get; set; } = 1;

        [JsonPropertyName("profiles")]
        public Dictionary<string, AuthProfileEntry> Profiles { get; set; } = new();

        [JsonPropertyName("usageStats")]
        public Dictionary<string, UsageStatsEntry> UsageStats { get; set; } = new();
    }

    public class AuthProfileEntry
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "api_key";

        [JsonPropertyName("provider")]
        public string Provider { get; set; } = "kimi-coding";

        [JsonPropertyName("key")]
        public string? Key { get; set; }
    }

    public class UsageStatsEntry
    {
        [JsonPropertyName("errorCount")]
        public int ErrorCount { get; set; } = 0;

        [JsonPropertyName("lastUsed")]
        public long LastUsed { get; set; }
    }
}
