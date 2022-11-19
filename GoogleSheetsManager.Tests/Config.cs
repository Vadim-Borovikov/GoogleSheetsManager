using System.Collections.Generic;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Tests;

internal sealed class Config
{
    [UsedImplicitly]
    public Dictionary<string, string?>? GoogleCredential { get; set; }
    [UsedImplicitly]
    public string? GoogleSheetId { get; set; }
}