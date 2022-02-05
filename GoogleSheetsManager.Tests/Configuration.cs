using System.Collections.Generic;
using Newtonsoft.Json;

namespace GoogleSheetsManager.Tests;

internal sealed class Configuration
{
    [JsonProperty]
    public Dictionary<string, string>? GoogleCredential { get; set; }

    [JsonProperty]
    public string? GoogleSheetId { get; set; }
}