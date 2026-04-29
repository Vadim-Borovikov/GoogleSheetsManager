using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;

namespace GoogleSheetsManager.Providers;

internal sealed class DriveProvider
{
    public DriveProvider(DriveService service, string fileId)
    {
        _service = service;
        _fileId = fileId;
    }

    public Task DeleteSpreadsheetAsync()
    {
        FilesResource.DeleteRequest request = new(_service, _fileId);
        return request.ExecuteAsync();
    }

    public async Task<string> GetNameAsync()
    {
        Google.Apis.Drive.v3.Data.File file = await GetFileAsync();
        return file.Name;
    }

    public async Task<IList<string>> GetParentsAsync()
    {
        Google.Apis.Drive.v3.Data.File file = await GetFileAsync("parents");
        return file.Parents;
    }

    public async Task MoveAndRenameAsync(string name, string folderId, string oldParents)
    {
        Google.Apis.Drive.v3.Data.File body = new() { Name = name };

        FilesResource.UpdateRequest request = _service.Files.Update(body, _fileId);
        request.AddParents = folderId;
        request.RemoveParents = oldParents;

        await request.ExecuteAsync();
    }

    public Task AddPermissionToAsync(string type, string role, string? emailAddress)
    {
        Permission body = new()
        {
            Type = type,
            Role = role,
            EmailAddress = emailAddress
        };

        PermissionsResource.CreateRequest request = _service.Permissions.Create(body, _fileId);

        return request.ExecuteAsync();
    }

    public Task DownloadAsync(string mimeType, Stream stream)
    {
        FilesResource.ExportRequest request = _service.Files.Export(_fileId, mimeType);
        return request.DownloadAsync(stream);
    }

    private Task<Google.Apis.Drive.v3.Data.File> GetFileAsync(string? fields = null)
    {
        FilesResource.GetRequest request = _service.Files.Get(_fileId);
        if (!string.IsNullOrWhiteSpace(fields))
        {
            request.Fields = fields;
        }
        return request.ExecuteAsync();
    }

    private readonly DriveService _service;
    private readonly string _fileId;
}