using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;

namespace GoogleSheetsManager.Providers;

internal sealed class DriveProvider : IDisposable
{
    public DriveProvider(BaseClientService.Initializer initializer, string fileId)
    {
        _service = new DriveService(initializer);
        _fileId = fileId;
    }

    public void Dispose() => _service.Dispose();

    public Task DeleteSpreadsheetAsync()
    {
        FilesResource.DeleteRequest request = new(_service, _fileId);
        return request.ExecuteAsync();
    }

    public async Task<IList<string>> GetParentsAsync()
    {
        File file = await GetFileAsync("parents");
        return file.Parents;
    }

    public async Task MoveAndRenameAsync(string name, string folderId, string oldParents)
    {
        File body = new() { Name = name };

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

    private Task<File> GetFileAsync(string fields)
    {
        FilesResource.GetRequest request = _service.Files.Get(_fileId);
        request.Fields = fields;
        return request.ExecuteAsync();
    }

    private readonly DriveService _service;
    private readonly string _fileId;
}