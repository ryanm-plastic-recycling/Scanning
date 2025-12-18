using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.SignalR;
using System.Text.Json; 

var builder = WebApplication.CreateBuilder(args);

// Let the server listen on specific IP addresses and hostnames.
builder.WebHost.UseUrls(
    "http://192.168.48.153:5026",
    "http://PRI-R-A03:5026",
    "http://localhost:5026");

// Ensure exact JSON names for APIs.
builder.Services.Configure<JsonOptions>(options =>
{
    options.SerializerOptions.PropertyNamingPolicy = null;
});

// Enable CORS if your SharePoint page is served from a different origin.
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSharePoint",
        policy => policy
            .WithOrigins(
                "https://plasticrecycling.sharepoint.com",
                "http://localhost:3000")
            .AllowAnyHeader()
            .AllowAnyMethod());
});

// Add SignalR support.
builder.Services.AddSignalR();

// Register the RFIDService with SignalR.
builder.Services.AddSingleton<RFIDService>(provider =>
    new RFIDService(provider.GetRequiredService<IHubContext<RFIDHub>>()));

var app = builder.Build();

// Use CORS.
app.UseCors("AllowSharePoint");

// Serve static files.
app.UseDefaultFiles();
app.UseStaticFiles();

var rfidService = app.Services.GetRequiredService<RFIDService>();

// Initialize the two readers with the new IPs and naming conventions:
// Primary (Reader56) uses 192.168.48.251 wifi address 192.168.49.24
// Secondary (Reader7B) uses 192.168.50.250 wifi address 192.168.50.117
rfidService.InitializeReader("192.168.49.24");
rfidService.InitializeThirdReader("192.168.50.117");
//rfidService.InitializeReader("192.168.48.251");
//rfidService.InitializeThirdReader("192.168.50.250");

// Map the SignalR hub at "/rfidHub".
app.MapHub<RFIDHub>("/rfidHub");

// API: Start Scanning.
app.MapPost("/api/reader/start", () =>
{
    try
    {
        rfidService.StartReading();
        return Results.Ok("RFID reading started");
    }
    catch (Exception ex)
    {
        return Results.Problem($"Failed to start RFID reader: {ex.Message}");
    }
});

// API: Stop Scanning.
app.MapPost("/api/reader/stop", () =>
{
    try
    {
        rfidService.StopReading();
        return Results.Ok("RFID reading stopped");
    }
    catch (Exception ex)
    {
        return Results.Problem($"Failed to stop RFID reader: {ex.Message}");
    }
});

// API: Clear tags for a specific dock door.
app.MapPost("/api/reader/clear/{dockDoor}", (int dockDoor) =>
{
    try
    {
        rfidService.ClearTags(dockDoor);
        return Results.Ok($"Tags cleared for Dock Door {dockDoor}");
    }
    catch (Exception ex)
    {
        return Results.Problem($"Failed to clear tags for Dock Door {dockDoor}: {ex.Message}");
    }
});

// API: Get real-time RFID tag data.
app.MapGet("/api/reader/tags", () =>
{
    try
    {
        var tagsByDock = rfidService.GetTagsByDock();
        return Results.Ok(tagsByDock);
    }
    catch (Exception ex)
    {
        return Results.Problem($"Error retrieving RFID tag data: {ex.Message}");
    }
});

// API: Export data to Excel per dock door.
app.MapPost("/api/reader/export/{dockDoor}/{billOfLading}", (int dockDoor, string billOfLading) =>
{
    try
    {
        var fileBytes = rfidService.CreateExcelFileBytes(dockDoor, billOfLading);
        return Results.File(
            fileBytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"Truck_{billOfLading}.xlsx"
        );
    }
    catch (Exception ex)
    {
        return Results.Problem($"Failed to export file: {ex.Message}");
    }
});

// API: Force refresh of the Excel cache (with a simple password check).
app.MapPost("/api/reader/refreshExcelCache", (HttpContext httpContext) =>
{
    if (!httpContext.Request.Headers.TryGetValue("X-Refresh-Password", out var password)
        || password != "SuperSecret123")
    {
        return Results.Unauthorized();
    }

    SharePointExcelService.ForceRefreshCache();
    return Results.Ok("Cache forced to refresh at " + DateTime.UtcNow);
});

// API: A simple endpoint to refresh data for the dashboard.
app.MapPost("/api/refreshData", (HttpContext httpContext) =>
{
    // Optionally validate a header or token here.
    SharePointExcelService.ForceRefreshCache();
    return Results.Ok("Data refreshed at " + DateTime.UtcNow);
});

// API: Preview the cached DB.xlsx data for diagnostics/UI verification.
app.MapGet("/api/db/preview", async () =>
{
    var entries = await SharePointExcelService.GetAllEntriesAsync();
    return Results.Ok(entries);
});

app.Run();
