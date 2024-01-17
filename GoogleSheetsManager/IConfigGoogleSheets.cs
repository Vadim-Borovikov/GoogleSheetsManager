using System.Collections.Generic;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public interface IConfigGoogleSheets
{
    public Dictionary<string, string> Credential { get; }

    public string ApplicationName { get; }

    public string? TimeZoneId { get; }
}