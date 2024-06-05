using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;

namespace MicrosoftGraphTest;

public static class GetAllUsersWithPagination
{
    /// <summary>
    /// Pagination documentation
    /// https://learn.microsoft.com/en-us/graph/paging?tabs=http
    /// 
    /// Remember to give the application permission to "User.Read.All" in Azure Portal and grant admin consent.
    /// </summary>
    public static async Task Run()
    {
        var clientId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_TENANT_ID");

        // We need any instance of TokenCredential to get the token.
        // Where we use InteractiveBrowserCredential to get the token interactively (browser).
        var credentialOptions = new InteractiveBrowserCredentialOptions { ClientId = clientId, TenantId = tenantId };
        var credential = new InteractiveBrowserCredential(credentialOptions);

        // It's not necessary to define any scope because they're already defined in the app registration.
        var token = credential.GetToken(new TokenRequestContext());

        var users = new List<JsonNode>();

        var client = new HttpClient();
        var url = "https://graph.microsoft.com/v1.0/users?$top=10";
        var nextLink = url;
        while (nextLink != null)
        {
            Console.WriteLine($"Sending request to {nextLink}");

            var request = new HttpRequestMessage(HttpMethod.Get, nextLink);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            var response = await client.SendAsync(request);

            var responseContentRaw = await response.Content.ReadAsStringAsync();
            var responseContent = JsonSerializer.Deserialize<JsonNode>(responseContentRaw);

            // Pretty print the JSON response.
            // Console.WriteLine(JsonSerializer.Serialize(responseContent, new JsonSerializerOptions { WriteIndented = true }));

            var value = responseContent?["value"]?.AsArray();
            if(value is null)
            {
                break;
            }

            foreach (var x in value)
            {
                users.Add(x);
            }

            nextLink = responseContent?["@odata.nextLink"]?.GetValue<string>();
        }

        Console.WriteLine($"Total users: {users.Count}");

        Console.ReadKey();
    }
}
