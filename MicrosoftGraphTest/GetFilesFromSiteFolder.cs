using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;

namespace MicrosoftGraphTest;

public static class GetFilesFromSiteFolder
{
    /// <summary>
    /// List all files in a site's one drive folder unders specific path.
    /// 
    /// Business case:
    /// We have a Team group (with corresponding SharePoint site), on OneDrive of this group we have a folder with *.xlsx files containing templates for offers.
    /// We want to list all files in this folder to allow user to choose one of them.
    /// 
    /// Required application (client) permissions:
    /// - Sites.Read.All
    /// </summary>
    public static async Task Run()
    {
        var clientId = Environment.GetEnvironmentVariable("GRAPH_API_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("GRAPH_API_TENANT_ID");
        var siteId = Environment.GetEnvironmentVariable("GRAPH_API_SITE_ID");
        var folderPath = "/oferty-szablony";

        // We need any instance of TokenCredential to get the token.
        // Where we use InteractiveBrowserCredential to get the token interactively (browser).
        var credentialOptions = new InteractiveBrowserCredentialOptions { ClientId = clientId, TenantId = tenantId };
        var credential = new InteractiveBrowserCredential(credentialOptions);

        // It's not necessary to define any scope because they're already defined in the app registration.
        var token = await credential.GetTokenAsync(new TokenRequestContext());

        var client = new HttpClient();

        // Please notice that path is wrapped in colons :{folderPath}:
        var url = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drive/root:{folderPath}:/children";
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        var response = await client.SendAsync(request);
        var responseContent = await response.Content.ReadAsStringAsync();

        Console.WriteLine(JsonSerializer.Serialize(responseContent, new JsonSerializerOptions { WriteIndented = true }));

        var items = JsonSerializer.Deserialize<JsonNode>(responseContent)?["value"]?.AsArray();

        foreach (var item in items)
        {
            Console.WriteLine($"Name: {item["name"]?.ToString()}");
        }

        Console.ReadKey();
    }
}
