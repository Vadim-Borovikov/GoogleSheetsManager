using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsManager.Extensions;
using GoogleSheetsManager.Providers;
using GryphonUtilities.Time;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Documents;

[PublicAPI]
public sealed class Manager : IDisposable
{
    public Manager(IConfigGoogleSheets config)
    {
        string credentialJson = JsonSerializer.Serialize(config.Credential);

        ServiceAccountCredential credential = CredentialFactory.FromJson<ServiceAccountCredential>(credentialJson);
        credential.Scopes = SheetsProvider.Scopes;
        _serviceInitializer = new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = config.ApplicationName
        };
        _sheetsService = new SheetsService(_serviceInitializer);
        _driveService = new DriveService(_serviceInitializer);

        _documents = new Dictionary<string, Document>();

        Clock clock = new(config.TimeZoneId);
        _converters[typeof(DateTimeFull)] = _converters[typeof(DateTimeFull?)] = o => o?.ToDateTimeFull(clock);
        _converters[typeof(DateOnly)] = _converters[typeof(DateOnly?)] = o => o?.ToDateOnly(clock);
        _converters[typeof(TimeOnly)] = _converters[typeof(TimeOnly?)] = o => o?.ToTimeOnly(clock);
        _converters[typeof(TimeSpan)] = _converters[typeof(TimeSpan?)] = o => o?.ToTimeSpan(clock);
    }

    public void Dispose()
    {
        _documents.Clear();
        _sheetsService.Dispose();
        _driveService.Dispose();
    }

    public Document GetOrAdd(string id)
    {
        if (!_documents.ContainsKey(id))
        {
            _documents[id] = new Document(_sheetsService, _driveService, id, _converters);
        }
        return _documents[id];
    }

    private Document GetOrAdd(Spreadsheet spreadsheet)
    {
        if (!_documents.ContainsKey(spreadsheet.SpreadsheetId))
        {
            _documents[spreadsheet.SpreadsheetId] =
                new Document(_sheetsService, _driveService, spreadsheet, _converters);
        }

        return _documents[spreadsheet.SpreadsheetId];
    }

    public async Task<Document> CopyAsync(string sourceId, string newName, string folderId)
    {
        Document from = GetOrAdd(sourceId);
        Spreadsheet fromSpreadsheet = await from.GetSpreadsheetAsync();
        Spreadsheet toSpreadsheet = await from.SheetsProvider.CreateNewSpreadsheetAsync(fromSpreadsheet.Properties);
        Document to = GetOrAdd(toSpreadsheet);
        to.SheetsProvider.PlanToDeleteSheets(toSpreadsheet);
        await to.SheetsProvider.CopyContentAndPlanToRenameSheetsAsync(from.SheetsProvider, fromSpreadsheet);
        await to.SheetsProvider.ExecutePlannedAsync();

        DriveProvider driveProvider = new(_driveService, toSpreadsheet.SpreadsheetId);
        await driveProvider.AddPermissionToAsync("anyone", "writer", null);
        await MoveAndRenameAsync(driveProvider, newName, folderId);

        return to;
    }

    public async Task DeleteAsync(string id)
    {
        DriveProvider driveProvider = new(_driveService, id);
        await driveProvider.DeleteSpreadsheetAsync();
        _documents.Remove(id);
    }

    public async Task DownloadAsync(string id, string mimeType, string path)
    {
        using (MemoryStream stream = new())
        {
            await DownloadAsync(id, mimeType, stream);
            await File.WriteAllBytesAsync(path, stream.ToArray());
        }
    }

    public async Task DownloadAsync(string id, string mimeType, Stream stream)
    {
        DriveProvider driveProvider = new(_driveService, id);
        await driveProvider.DownloadAsync(mimeType, stream);
    }

    private static async Task MoveAndRenameAsync(DriveProvider provider, string newName, string folderId)
    {
        IList<string> oldParentLists = await provider.GetParentsAsync();
        string oldParents = string.Join(',', oldParentLists);
        await provider.MoveAndRenameAsync(newName, folderId, oldParents);
    }

    private readonly BaseClientService.Initializer _serviceInitializer;
    private readonly SheetsService _sheetsService;
    private readonly DriveService _driveService;
    private readonly IDictionary<string, Document> _documents;

    private readonly Dictionary<Type, Func<object?, object?>> _converters = new()
    {
        { typeof(bool), v => v.ToBool() },
        { typeof(bool?), v => v.ToBool() },
        { typeof(int), v => v.ToInt() },
        { typeof(int?), v => v.ToInt() },
        { typeof(byte), v => v.ToByte() },
        { typeof(byte?), v => v.ToByte() },
        { typeof(long), v => v.ToLong() },
        { typeof(long?), v => v.ToLong() },
        { typeof(decimal), v => v.ToDecimal() },
        { typeof(decimal?), v => v.ToDecimal() },
        { typeof(string), v => v?.ToString() },
    };
}