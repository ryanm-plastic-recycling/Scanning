using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;

public class SharePointExcelService
{
    // Cache for Excel rows and lookup dictionary.
    private static JArray? cachedRows;
    private static Dictionary<string, JObject> cachedLookup = new Dictionary<string, JObject>(StringComparer.OrdinalIgnoreCase);
    private static DateTime cacheExpiration = DateTime.MinValue;
    private static readonly TimeSpan cacheDuration = TimeSpan.FromHours(1);
    private const string WorkbookFileName = "DB.xlsx";
    private const string WorksheetName = "DB";
    private const string TableName = "DBRFID";

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
    /// Returns a snapshot of all cached rows for diagnostics/UI preview.
    /// </summary>
    public static async Task<JArray> GetAllEntriesAsync()
    {
        if (DateTime.UtcNow >= cacheExpiration || cachedRows == null)
        {
            await LoadCacheAsync();
        }

        var rows = new JArray();
        foreach (var kvp in cachedLookup)
        {
            var entry = new JObject
            {
                { "RfidBox", kvp.Key }
            };

            foreach (var property in kvp.Value)
            {
                entry[property.Key] = property.Value;
            }

            rows.Add(entry);
        }

        return rows;
    }

    /// <summary>
    /// Loads the Excel table from the local DB.xlsx file and builds a lookup dictionary.
    /// </summary>
    private static Task LoadCacheAsync()
    {
        Console.WriteLine("Loading Excel data from local DB.xlsx...");

        string excelPath = Path.Combine(AppContext.BaseDirectory, WorkbookFileName);
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Excel file not found at {excelPath}");
            cachedRows = new JArray();
            cachedLookup.Clear();
            cacheExpiration = DateTime.MinValue;
            return Task.CompletedTask;
        }

        try
        {
            using var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                ws.Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet == null)
            {
                Console.WriteLine($"Worksheet '{WorksheetName}' not found in {WorkbookFileName}.");
                cachedRows = new JArray();
                cachedLookup.Clear();
                cacheExpiration = DateTime.MinValue;
                return Task.CompletedTask;
            }

            var table = worksheet.Tables
                .FirstOrDefault(t => t.Name.Equals(TableName, StringComparison.OrdinalIgnoreCase));
            if (table == null)
            {
                Console.WriteLine($"Table '{TableName}' not found in worksheet '{WorksheetName}'.");
                cachedRows = new JArray();
                cachedLookup.Clear();
                cacheExpiration = DateTime.MinValue;
                return Task.CompletedTask;
            }

            var headerLookup = table.Fields
                .Select((field, index) => new { field.Name, Index = index })
                .ToDictionary(
                    f => f.Name.Trim(),
                    f => f.Index,
                    StringComparer.OrdinalIgnoreCase);

            cachedLookup.Clear();

            foreach (var row in table.DataRange.Rows())
            {
                string fileRfid = GetCellValue(row, headerLookup, "RFIDBOX");
                if (string.IsNullOrWhiteSpace(fileRfid))
                    continue;

                JObject extraData = new JObject
                {
                    { "Lot", GetCellValue(row, headerLookup, "RFID", "N/A") },
                    { "Dept", GetCellValue(row, headerLookup, "DEPARTMENT", "N/A") },
                    { "Row", GetCellValue(row, headerLookup, "ROW", "N/A") },
                    { "DeptLot", GetCellValue(row, headerLookup, "DEPARTMENT LOT", "N/A") },
                    { "Supplier", GetCellValue(row, headerLookup, "CUSTOMER", "N/A") },
                    { "Type", GetCellValue(row, headerLookup, "TYPE", "N/A") },
                    { "Color", GetCellValue(row, headerLookup, "COLOR", "N/A") },
                    { "Format", GetCellValue(row, headerLookup, "FORMAT", "N/A") },
                    { "Pounds", GetCellValue(row, headerLookup, "POUNDS", "N/A") },
                    { "Price", GetCellValue(row, headerLookup, "PRICE", "N/A") },
                    { "Freight", GetCellValue(row, headerLookup, "FREIGHT", "N/A") },
                    { "Toll", GetCellValue(row, headerLookup, "TOLLING", "N/A") },
                    { "Date", GetCellValue(row, headerLookup, "Date", "N/A") }
                };

                cachedLookup[fileRfid.Trim()] = extraData;
            }

            cachedRows = new JArray();
            cacheExpiration = DateTime.UtcNow.Add(cacheDuration);
            Console.WriteLine($"Finished building local dictionary for RFID lookups. {cachedLookup.Count} entries loaded.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading Excel data: {ex.Message}");
            cachedRows = new JArray();
            cachedLookup.Clear();
            cacheExpiration = DateTime.MinValue;
        }

        return Task.CompletedTask;
    }

    private static string GetCellValue(IXLRangeRow row, Dictionary<string, int> headerLookup, string columnName, string fallback = "")
    {
        if (headerLookup.TryGetValue(columnName, out int index))
        {
            // ClosedXML rows are 1-based for cells.
            string value = row.Cell(index + 1).GetString().Trim();
            if (!string.IsNullOrWhiteSpace(value))
                return value;
        }

        return fallback;
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
