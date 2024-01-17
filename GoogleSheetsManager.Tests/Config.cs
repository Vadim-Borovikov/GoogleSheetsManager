using System.Collections.Generic;
using JetBrains.Annotations;

// ReSharper disable NullableWarningSuppressionIsUsed

namespace GoogleSheetsManager.Tests;

internal sealed class Config : IConfigGoogleSheets
{
    [UsedImplicitly]
    public Dictionary<string, string> GoogleCredential { get; init; } = null!;

    public Dictionary<string, string> Credential => GoogleCredential;
    public string ApplicationName { get; init; } = null!;
    public string? TimeZoneId { get; init; }

    public string? GoogleSheetId { get; init; }
}