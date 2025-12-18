using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;

public class SharePointExcelService
{
    private static HttpClient _httpClient = new HttpClient();

    // Cache for Excel rows and lookup dictionary.
    private static JArray? cachedRows;
    private static Dictionary<string, JObject> cachedLookup = new Dictionary<string, JObject>();
    private static DateTime cacheExpiration = DateTime.MinValue;
    private static readonly TimeSpan cacheDuration = TimeSpan.FromHours(1);

    /// <summary>
    /// Returns extra data for the specified RFID tag from the local cache.
    /// </summary>
    public static async Task<JObject?> GetExtraDataForTagLocal(string rfidTag)
    {
        if (DateTime.UtcNow >= cacheExpiration || cachedRows == null)
        {
            await LoadCacheAsync();
        }

        if (cachedLookup == null || cachedLookup.Count == 0)
            return null;

        cachedLookup.TryGetValue(rfidTag.Trim(), out JObject? extraData);
        return extraData;
    }

    /// <summary>
    /// Loads the Excel table from Graph and builds a lookup dictionary.
    /// </summary>
    private static async Task LoadCacheAsync()
    {
        Console.WriteLine("Loading Excel data from Graph...");
        string graphUrl = "https://graph.microsoft.com/v1.0/drives/b!LWoyGNela0icnELnsn9fq_T-dScYDjNGoowEPLikhf6Ug4zYbAbTQrP5_r8NC_0k/items/013D6T55JX7EPIKX35FJAI5DKFANSQVDBO/workbook/tables/DBRFID/rows?$select=index,values";
        string token = await GraphAuthHelper.GetAccessTokenAsync();
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

        var response = await _httpClient.GetAsync(graphUrl);
        response.EnsureSuccessStatusCode();

        var content = await response.Content.ReadAsStringAsync();
        JObject data = JObject.Parse(content);

        var valueToken = data["value"];
        if (valueToken is JArray rows)
        {
            cachedRows = rows;
            Console.WriteLine($"Loaded {rows.Count} rows into cache.");
        }
        else
        {
            cachedRows = new JArray();
            Console.WriteLine("No rows returned from Excel.");
        }

        cachedLookup.Clear();
        if (cachedRows != null)
        {
            int rfidBoxColumnIndex = 18; // Column S (0-indexed)
            foreach (var row in cachedRows)
            {
                var valuesToken = row["values"];
                if (valuesToken == null || !valuesToken.HasValues)
                    continue;

                if (valuesToken[0] is JArray rowValues && rowValues.Count > rfidBoxColumnIndex)
                {
                    string fileRfid = rowValues[rfidBoxColumnIndex]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(fileRfid))
                    {
                        JObject extraData = new JObject
                        {
                            { "Lot", rowValues.Count > 0 ? rowValues[0]?.ToString().Trim() ?? "N/A" : "N/A" },
                            { "Type", rowValues.Count > 9 ? rowValues[9]?.ToString().Trim() ?? "N/A" : "N/A" },
                            { "Color", rowValues.Count > 10 ? rowValues[10]?.ToString().Trim() ?? "N/A" : "N/A" },
                            { "Format", rowValues.Count > 11 ? rowValues[11]?.ToString().Trim() ?? "N/A" : "N/A" },
                            { "Pounds", rowValues.Count > 14 ? rowValues[14]?.ToString().Trim() ?? "N/A" : "N/A" }
                        };

                        cachedLookup[fileRfid] = extraData;
                    }
                }
            }
        }

        cacheExpiration = DateTime.UtcNow.Add(cacheDuration);
        Console.WriteLine("Finished building local dictionary for RFID lookups.");
    }

    /// <summary>
    /// Force a cache refresh.
    /// </summary>
    public static void ForceRefreshCache()
    {
        cachedRows = null;
        cachedLookup.Clear();
        cacheExpiration = DateTime.MinValue;
        Console.WriteLine("Cache forced to refresh.");
    }
}
