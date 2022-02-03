using System;
using System.Collections.Generic;
using System.Linq;
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

    public async Task<IEnumerable<string>> GetPermissionIdsAsync()
    {
        PermissionList permissions = await GetPermissionsAsync();
        return permissions.Permissions.Select(p => p.Id);
    }

    public Task DowngradePermissionAsync(string id) => UpdatePermissionAsync(id, "reader");

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

    public Task AddPermissionToAsync(string type, string role, string emailAddress, bool transferOwnership)
    {
        Permission body = new()
        {
            Type = type,
            Role = role,
            EmailAddress = emailAddress
        };

        PermissionsResource.CreateRequest request = _service.Permissions.Create(body, _fileId);
        request.TransferOwnership = transferOwnership;

        return request.ExecuteAsync();
    }

    private Task<PermissionList> GetPermissionsAsync()
    {
        PermissionsResource.ListRequest request = _service.Permissions.List(_fileId);
        return request.ExecuteAsync();
    }

    private Task UpdatePermissionAsync(string id, string role)
    {
        Permission body = new()
        {
            Role = role
        };
        PermissionsResource.UpdateRequest request = _service.Permissions.Update(body, _fileId, id);
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