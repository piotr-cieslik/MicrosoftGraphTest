using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;

namespace MicrosoftGraphTest;

public static class ReplaceValuesInExcelFile
{
    /// <summary>
    /// This script is a part of a bigger project where we want to create a new offer based on a template.
    /// Steps pefromed by this script:
    /// 1. Find a template file in a specific folder.
    /// 2. Ask user to choose a template file.
    /// 3. Create a copy of the selected template file in a target folder.
    /// 4. Replace values in table.
    /// 
    /// Business case:
    /// We have a Team group (with corresponding SharePoint site), on OneDrive of this group we have a folder with *.xlsx files containing templates for offers.
    /// We want to list all files in this folder to allow user to choose one of them.
    /// 
    /// Required application (client) permissions:
    /// - Sites.ReadWrite.All
    /// </summary>
    public static async Task Run()
    {
        var clientId = Environment.GetEnvironmentVariable("GRAPH_API_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("GRAPH_API_TENANT_ID");
        var siteId = Environment.GetEnvironmentVariable("GRAPH_API_SITE_ID");
        var templateFolderPath = "/oferty-szablony";
        var targetFolderPath = "/oferty";

        // We need any instance of TokenCredential to get the token.
        // Where we use InteractiveBrowserCredential to get the token interactively (browser).
        var credentialOptions = new InteractiveBrowserCredentialOptions { ClientId = clientId, TenantId = tenantId };
        var credential = new InteractiveBrowserCredential(credentialOptions);

        // It's not necessary to define any scope because they're already defined in the app registration.
        var token = await credential.GetTokenAsync(new TokenRequestContext());
        Console.WriteLine($"Token: {token.Token}");

        // We need to use HttpClient to send requests to the Graph API.
        var client = new HttpClient();

        var driveId = await GetDriveIdOfSite(client, token);

        // Please notice that path is wrapped in colons :{folderPath}:
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:{templateFolderPath}:/children";
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        var response = await client.SendAsync(request);
        var responseContent = await response.Content.ReadAsStringAsync();

        var templateFiles = JsonSerializer.Deserialize<JsonNode>(responseContent)?["value"]?.AsArray();
        var templateFileId = templateFiles?[0]?["id"]?.ToString();
        var targetFileId = await CreateCopyOfTemplateFile(client, token, driveId, templateFileId, targetFolderPath);

        await ReplaceVariables(client, token, driveId, targetFileId);

        Console.ReadLine();
    }

    public static async Task<string> GetDriveIdOfSite(HttpClient client, AccessToken token)
    {
        var siteId = Environment.GetEnvironmentVariable("GRAPH_API_SITE_ID");

        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive";
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        var response = await client.SendAsync(request);
        var responseContent = await response.Content.ReadAsStringAsync();

        var driveId = JsonSerializer.Deserialize<JsonNode>(responseContent)?["id"]?.ToString();

        if (driveId is null)
        {
            throw new Exception("Drive ID not found.");
        }

        Console.WriteLine($"Drive ID: {driveId}");
        return driveId;
    }

    public static async Task<string> GetFolderIdOfDrive(HttpClient client, AccessToken token, string driveId, string folderPath)
    {
        var url = $"https://graph.microsoft.com/v1.0/drives/{driveId}/root:{folderPath}:/";
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        var response = await client.SendAsync(request);
        var responseContent = await response.Content.ReadAsStringAsync();

        var folderId = JsonSerializer.Deserialize<JsonNode>(responseContent)?["id"]?.ToString();

        if (folderId is null)
        {
            throw new Exception("Folder ID not found.");
        }

        Console.WriteLine($"Folder ID: {folderId}");
        return folderId;
    }

    /// <summary>
    /// Create a copy of a selected template file in the target folder.
    /// </summary>
    /// <returns>Copy item (file) id.</returns>
    public static async Task<string> CreateCopyOfTemplateFile(HttpClient client, AccessToken token, string driveId, string templateFileId, string targetFolderPath)
    {
        var targetFolderId = await GetFolderIdOfDrive(client, token, driveId, targetFolderPath);
        var targetFileName = $"{DateTime.UtcNow.Ticks}.xlsx";

        // Copy it to the target folder
        // In theory you can use Path property instead of Id, but it's not working for me.
        // https://learn.microsoft.com/en-us/graph/api/resources/itemreference?view=graph-rest-1.0
        var copyUrl = $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{templateFileId}/copy";
        var copyRequest = new HttpRequestMessage(HttpMethod.Post, copyUrl);
        copyRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        copyRequest.Content = JsonContent.Create(new
        {
            ParentReference = new
            {
                DriveId = driveId,
                Id = targetFolderId,
            },
            Name = targetFileName,
        });

        var response = await client.SendAsync(copyRequest);
        response.EnsureSuccessStatusCode();

        var targetFileUrl = $"https://graph.microsoft.com/v1.0/drives/{driveId}/root:{targetFolderPath}/{targetFileName}";
        var targetFileRequest = new HttpRequestMessage(HttpMethod.Get, targetFileUrl);
        targetFileRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        var targetFileResponse = await client.SendAsync(targetFileRequest);
        targetFileResponse.EnsureSuccessStatusCode();
        var targetFileResponseContent = await targetFileResponse.Content.ReadAsStringAsync();

        var targetFile = JsonSerializer.Deserialize<JsonNode>(targetFileResponseContent);
        var targetFileId = targetFile?["id"]?.ToString();
        return targetFileId;
    }

    public static async Task ReplaceVariables(HttpClient client, AccessToken token, string driveId, string fileItemId)
    {
        var url = $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{fileItemId}/workbook/tables/ZmienneTabela/range?$select=values";
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        var response = await client.SendAsync(request);
        var responseContent = await response.Content.ReadAsStringAsync();
        var rows = JsonSerializer.Deserialize<JsonNode>(responseContent)?["values"]?.AsArray();

        var updatedRows = new List<List<string>>();
        foreach (var row in rows)
        {
            var updatedRow = new List<string>();
            foreach(var cell in row.AsArray())
            {
                var cellValue = cell.ToString();
                if (cellValue.Contains("."))
                {
                    cellValue = cellValue.Replace(".", "...");
                }
                updatedRow.Add(cellValue);
            }
            updatedRows.Add(updatedRow);
        }

        var updateUrl = $"https://graph.microsoft.com/v1.0/drives/{driveId}/items/{fileItemId}/workbook/tables/ZmienneTabela/range";
        var updateRequest = new HttpRequestMessage(HttpMethod.Patch, updateUrl);
        updateRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        updateRequest.Content = JsonContent.Create(new
        {
            values = updatedRows,
        });

        var updateResponse = await client.SendAsync(updateRequest);
        updateResponse.EnsureSuccessStatusCode();
    }
}