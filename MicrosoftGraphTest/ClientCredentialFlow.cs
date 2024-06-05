using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;

namespace MicrosoftGraphTest;

public static class ClientCredentialFlow
{
    public static async Task Run()
    {
        var clientId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_TENANT_ID");
        var clientSecret = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_CLIENT_SECRET");

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

        // The .default scope is used to get the token for the whole API.
        // It's special scope defined automatically for each app registration.
        // Keep in mind that application uses different set of permissions than user.
        // In azure we have to type of permissions: delegated and application.
        var token = await credential.GetTokenAsync(new TokenRequestContext([$"api://{clientId}/.default"]));

        var client = new HttpClient();

        var url = "https://graph.microsoft.com/v1.0/users?$top=10";
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        var response = await client.SendAsync(request);

        var responseContentRaw = await response.Content.ReadAsStringAsync();
        Console.WriteLine(JsonSerializer.Serialize(JsonSerializer.Deserialize<JsonNode>(responseContentRaw), new JsonSerializerOptions { WriteIndented = true }));
    }
}
