using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;

namespace MicrosoftGraphTest;

public static class FetchBasicUserInformation
{
    public static async Task Run()
    {
        // After strugling a lot with MS Graph SDK for .NET, I decided to use HttpClient directly.
        // It's also what is recommended in this great article:
        // https://laurakokkarinen.com/the-ultimate-beginners-guide-to-microsoft-graph/

        var clientId = Environment.GetEnvironmentVariable("GRAPH_API_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("GRAPH_API_TENANT_ID");

        // We need any instance of TokenCredential to get the token.
        // Where we use InteractiveBrowserCredential to get the token interactively (browser).
        var credentialOptions = new InteractiveBrowserCredentialOptions { ClientId = clientId, TenantId = tenantId };
        var credential = new InteractiveBrowserCredential(credentialOptions);

        // It's not necessary to define any scope because they're already defined in the app registration.
        var token = credential.GetToken(new TokenRequestContext());

        var client = new HttpClient();

        // We need to set the Authorization header with the token for each request.
        var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        var response = await client.SendAsync(request);

        var responseContentRaw = await response.Content.ReadAsStringAsync();
        Console.WriteLine(responseContentRaw);

        var responseContent = JsonSerializer.Deserialize<JsonNode>(responseContentRaw);

        // Access property via indexer, then call GetValue<T> to convert to the desired type.
        var id = responseContent["id"].GetValue<Guid>();

        Console.WriteLine();
    }
}
