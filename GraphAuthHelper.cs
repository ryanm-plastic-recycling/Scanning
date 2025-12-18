using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

public class GraphAuthHelper
{
    private static readonly string clientId = "3c1a7c51-51a7-4cea-b847-cdc8eb2302fb";
    private static readonly string clientSecret = "GT-8Q~QodnuqvdTqEtvef2iPFlyU1glZz1e8Pbpa";
    private static readonly string tenantId = "ca800f2c-47b3-4400-8eb1-fb1db2a39a1e";

    private static IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

    private static readonly string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

    public static async Task<string> GetAccessTokenAsync()
    {
        try
        {
            var authResult = await clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
            return authResult.AccessToken;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error acquiring token: {ex.Message}");
            throw;
        }
    }
}
