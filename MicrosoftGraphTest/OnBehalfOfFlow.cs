using Azure.Core;
using Azure.Identity;

namespace MicrosoftGraphTest;

public static class OnBehalfOfFlow
{
    /// <summary>
    /// Example of On-Behalf-Of flow usages.
    /// In this flow we have 3 parties:
    /// 1. SPA (Single Page Application) - user facing application.
    /// 2. API - application that is called by SPA.
    /// 3. Graph API - application that is called by API.
    /// 
    /// Because access token is created for specific audience (resource), we need to use OnBehalfOfCredential to get the token for Graph API.
    /// 
    /// Please notice that you need to define scopes when requesting a token to ensure that you get correct one.
    /// There is a difference between access token for Graph API and access token for custom API.
    /// </summary>
    public static async Task Run()
    {
        // First obtain the access token as SPA to make a call to API.
        // My API is exposed under single scope "default". Make sure to use this scope when requesting the token.
        // If you don't define the scope, you'll get the access token for Graph API.

        var spaClientId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_CLIENT_ID");
        var spaTenantId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_SPA_TENANT_ID");

        var interactiveCredentialOptions = new InteractiveBrowserCredentialOptions { ClientId = spaClientId, TenantId = spaTenantId };
        var interactiveCredential = new InteractiveBrowserCredential(interactiveCredentialOptions);

        var apiAccessToken = await interactiveCredential.GetTokenAsync(new TokenRequestContext(["default"]));
        Console.WriteLine($"API access token: {apiAccessToken.Token}");

        // Then use the access token to get the token for Graph API.
        // You cannot use the API access token for Graph API because it's created for different audience.
        // We need to use on-behalf-of flow to exchange the API access token for Graph API access token.
        // 
        // Use '.default' as the scope to get the access token for Graph API.
        // Intead of '.default' you can use any other scope you defined in the app registration for Graph API.
        var apiClientId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_API_CLIENT_ID");
        var apiTenantId = Environment.GetEnvironmentVariable("GRAPH_API_TEST_API_TENANT_ID");
        var apiClientSecret = Environment.GetEnvironmentVariable("GRAPH_API_TEST_API_CLIENT_SECRET");

        var oboCredential = new OnBehalfOfCredential(apiTenantId, apiClientId, apiClientSecret, apiAccessToken.Token);

        using var oboCancellationTokenSource = new CancellationTokenSource();
        var graphAccessToken = await oboCredential.GetTokenAsync(new TokenRequestContext([".default"]), oboCancellationTokenSource.Token);
        Console.WriteLine($"Graph access token: {graphAccessToken.Token}");

        Console.ReadKey();
    }
}
