using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace OpenClawSetting.Models
{
    public class ModelsConfig
    {
        [JsonPropertyName("providers")]
        public Dictionary<string, ProviderConfig> Providers { get; set; } = new();
    }

    public class ProviderConfig
    {
        [JsonPropertyName("baseUrl")]
        public string BaseUrl { get; set; } = "https://api.kimi.com/coding/";

        [JsonPropertyName("api")]
        public string Api { get; set; } = "anthropic-messages";

        [JsonPropertyName("models")]
        public List<ModelEntry> Models { get; set; } = new();

        [JsonPropertyName("apiKey")]
        public string? ApiKey { get; set; }
    }

    public class ModelEntry
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = "k2p5";

        [JsonPropertyName("name")]
        public string Name { get; set; } = "Kimi for Coding";

        [JsonPropertyName("reasoning")]
        public bool Reasoning { get; set; } = true;

        [JsonPropertyName("input")]
        public List<string> Input { get; set; } = new List<string> { "text", "image" };

        [JsonPropertyName("cost")]
        public CostConfig Cost { get; set; } = new CostConfig();

        [JsonPropertyName("contextWindow")]
        public int ContextWindow { get; set; } = 262144;

        [JsonPropertyName("maxTokens")]
        public int MaxTokens { get; set; } = 32768;

        // 注意：JSON 里还有一个 "api" 字段，如果是重复的可以不加，加上也没错
        [JsonPropertyName("api")]
        public string Api { get; set; } = "anthropic-messages";
    }

    public class CostConfig
    {
        [JsonPropertyName("input")]
        public int Input { get; set; } = 0;

        [JsonPropertyName("output")]
        public int Output { get; set; } = 0;

        [JsonPropertyName("cacheRead")]
        public int CacheRead { get; set; } = 0;

        [JsonPropertyName("cacheWrite")]
        public int CacheWrite { get; set; } = 0;
    }
}
