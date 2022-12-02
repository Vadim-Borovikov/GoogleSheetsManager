using System.Collections.Generic;
using System.Text.Json;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public interface IConfigGoogleSheets
{
    public Dictionary<string, string>? Credential { get; }

    public string? CredentialJson { get; }

    public string ApplicationName { get; }

    public string? TimeZoneId { get; }

    public string GetCredentialJson()
    {
        return string.IsNullOrWhiteSpace(CredentialJson) ? JsonSerializer.Serialize(Credential) : CredentialJson;
    }
}