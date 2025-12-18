using Symbol.RFID3;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using Microsoft.AspNetCore.SignalR;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Http.Headers;

public class RFIDService
{
    // Primary reader ("Reader56") for Dock Doors 5 & 6.
    private RFIDReader? _reader;
    // Secondary reader ("Reader7B") for Dock Doors 7 & Bulk Load (8).
    private RFIDReader? _reader3;
    private readonly IHubContext<RFIDHub> _hubContext;
    private bool _isReading;
    private bool _isReading3;

    // Global scanned counter to preserve scanning order.
    private static int _scannedCounter = 0;

    // ConcurrentDictionary for storing all tags.
    // (Keys will include a suffix to differentiate readers.)
    private ConcurrentDictionary<string, dynamic> _tags = new ConcurrentDictionary<string, dynamic>();

    // Static HttpClient for any Graph calls.
    private static readonly HttpClient _httpClient = new HttpClient();

    public RFIDService(IHubContext<RFIDHub> hubContext)
    {
        _hubContext = hubContext;
    }

    // ---------------------------
    // READER CONFIGURATION
    // ---------------------------
    private void ConfigureReader(RFIDReader reader)
    {
        if (reader == null || !reader.IsConnected)
        {
            Console.WriteLine("Reader not connected or null in ConfigureReader");
            return;
        }

        try
        {
            // Configure antennas: set transmit power index and sensitivity.
            for (ushort ant = 0; ant < reader.Config.Antennas.Length; ant++)
            {
                var antConfig = reader.Config.Antennas[ant].GetRfConfig();
                antConfig.TransmitPowerIndex = 200; // 20 dBm
                antConfig.ReceiveSensitivityIndex = 0;
                reader.Config.Antennas[ant].SetRfConfig(antConfig);
            }

            // Optional: Configure singulation control for the first antenna.
            ushort antennaId = 0;
            var singControl = reader.Config.Antennas[antennaId].GetSingulationControl();
            singControl.Session = SESSION.SESSION_S1;
            singControl.TagPopulation = 30;
            reader.Config.Antennas[antennaId].SetSingulationControl(singControl);

            Console.WriteLine($"Reader {reader.HostName} configured.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error in ConfigureReader: " + ex.Message);
        }
    }

    // ---------------------------
    // INITIALIZATION
    // ---------------------------
    public void InitializeReader(string ip)
    {
        if (_reader == null)
        {
            _reader = new RFIDReader(ip, 5084, 0);
            _reader.Connect();

            ConfigureReader(_reader);

            _reader.Events.AttachTagDataWithReadEvent = true;
            _reader.Events.ReadNotify += OnTagRead;
            Console.WriteLine($"Initialized Reader56 at {ip}");
        }
    }

    public void InitializeThirdReader(string ip)
    {
        if (_reader3 == null)
        {
            _reader3 = new RFIDReader(ip, 5084, 0);
            _reader3.Connect();

            ConfigureReader(_reader3);

            _reader3.Events.AttachTagDataWithReadEvent = true;
            _reader3.Events.ReadNotify += OnThirdReaderTagRead;
            Console.WriteLine($"Initialized Reader7B at {ip}");
        }
    }

    // ---------------------------
    // START/STOP
    // ---------------------------
    public void StartReading()
    {
        if (!_isReading && _reader != null)
        {
            _reader.Actions.Inventory.Perform();
            _isReading = true;
            Console.WriteLine("Started inventory on Reader56.");
        }
        if (!_isReading3 && _reader3 != null)
        {
            _reader3.Actions.Inventory.Perform();
            _isReading3 = true;
            Console.WriteLine("Started inventory on Reader7B.");
        }
    }

    public void StopReading()
    {
        if (_isReading && _reader != null)
        {
            _reader.Actions.Inventory.Stop();
            _isReading = false;
            Console.WriteLine("Stopped inventory on Reader56.");
        }
        if (_isReading3 && _reader3 != null)
        {
            _reader3.Actions.Inventory.Stop();
            _isReading3 = false;
            Console.WriteLine("Stopped inventory on Reader7B.");
        }
    }

    public void ClearTags(int dockDoor)
    {
        if (dockDoor == 5 || dockDoor == 6)
        {
            foreach (var key in _tags.Keys.ToList())
            {
                if (key.EndsWith("_R56"))
                {
                    dynamic tag = _tags[key];
                    if (dockDoor == 5 && tag.AntennaID >= 1 && tag.AntennaID <= 4)
                        _tags.TryRemove(key, out _);
                    else if (dockDoor == 6 && tag.AntennaID >= 5 && tag.AntennaID <= 8)
                        _tags.TryRemove(key, out _);
                }
            }
        }
        else if (dockDoor == 7 || dockDoor == 8)
        {
            foreach (var key in _tags.Keys.ToList())
            {
                if (key.EndsWith("_R7B"))
                {
                    dynamic tag = _tags[key];
                    if (dockDoor == 7 && tag.AntennaID >= 1 && tag.AntennaID <= 4)
                        _tags.TryRemove(key, out _);
                    else if (dockDoor == 8 && tag.AntennaID >= 5 && tag.AntennaID <= 8)
                        _tags.TryRemove(key, out _);
                }
            }
        }
        Console.WriteLine($"Cleared tags for Dock Door {dockDoor}");
    }

    // ---------------------------
    // EVENT HANDLERS
    // ---------------------------
    private async void OnTagRead(object sender, Events.ReadEventArgs e)
    {
        try
        {
            var newTags = _reader?.Actions.GetReadTags(1000);
            if (newTags != null)
            {
                foreach (var t in newTags)
                {
                    string epcAscii = HexStringToAscii(t.TagID);
                    if (!(epcAscii.Contains("FF") || epcAscii.Contains("PL") ||
                          epcAscii.Contains("GL") || epcAscii.Contains("BL")))
                    {
                        continue;
                    }

                    int currentIndex = System.Threading.Interlocked.Increment(ref _scannedCounter);
                    JObject? extraData = await SharePointExcelService.GetExtraDataForTagLocal(epcAscii);
                    Console.WriteLine("Extra data for tag " + epcAscii + ": " +
                        (extraData != null ? extraData.ToString() : "null"));

                    var tagData = new
                    {
                        ScannedIndex = currentIndex,
                        EPC_Hex = t.TagID,
                        EPC_Ascii = epcAscii,
                        t.AntennaID,
                        t.PeakRSSI,
                        Lot = extraData?["Lot"]?.ToString() ?? "N/A",
                        Type = extraData?["Type"]?.ToString() ?? "N/A",
                        Color = extraData?["Color"]?.ToString() ?? "N/A",
                        Format = extraData?["Format"]?.ToString() ?? "N/A",
                        Pounds = extraData?["Pounds"]?.ToString() ?? "N/A",
                        Reader = "Primary"
                    };

                    _tags[t.TagID + "_R56"] = tagData;
                    await _hubContext.Clients.All.SendAsync("ReceiveTag", tagData);
                    Console.WriteLine($"Broadcast tag from Reader56: {tagData.EPC_Ascii}, idx={currentIndex}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OnTagRead: {ex.Message}");
        }
    }

    private async void OnThirdReaderTagRead(object sender, Events.ReadEventArgs e)
    {
        try
        {
            var newTags = _reader3?.Actions.GetReadTags(1000);
            if (newTags != null)
            {
                foreach (var t in newTags)
                {
                    string epcAscii = HexStringToAscii(t.TagID);
                    if (!(epcAscii.Contains("FF") || epcAscii.Contains("PL") ||
                          epcAscii.Contains("GL") || epcAscii.Contains("BL")))
                    {
                        continue;
                    }

                    int currentIndex = System.Threading.Interlocked.Increment(ref _scannedCounter);
                    JObject? extraData = await SharePointExcelService.GetExtraDataForTagLocal(epcAscii);

                    var tagData = new
                    {
                        ScannedIndex = currentIndex,
                        EPC_Hex = t.TagID,
                        EPC_Ascii = epcAscii,
                        t.AntennaID,
                        t.PeakRSSI,
                        Lot = extraData?["Lot"]?.ToString() ?? "N/A",
                        Type = extraData?["Type"]?.ToString() ?? "N/A",
                        Color = extraData?["Color"]?.ToString() ?? "N/A",
                        Format = extraData?["Format"]?.ToString() ?? "N/A",
                        Pounds = extraData?["Pounds"]?.ToString() ?? "N/A",
                        Reader = "Third"
                    };

                    _tags[t.TagID + "_R7B"] = tagData;
                    await _hubContext.Clients.All.SendAsync("ReceiveTag", tagData);
                    Console.WriteLine($"Broadcast tag from Reader7B: {tagData.EPC_Ascii}, idx={currentIndex}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in OnThirdReaderTagRead: {ex.Message}");
        }
    }

    // ---------------------------
    // UTILITY METHODS
    // ---------------------------
    public string HexStringToAscii(string hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
            return "";
        if (hex.Length % 2 != 0)
            return "<Invalid Hex Length>";
        try
        {
            hex = hex.Trim().Replace("\0", "").Replace(" ", "");
            byte[] rawBytes = new byte[hex.Length / 2];
            for (int i = 0; i < rawBytes.Length; i++)
            {
                rawBytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
            }
            string asciiString = Encoding.ASCII.GetString(rawBytes);
            asciiString = new string(asciiString.Where(c => c >= 32 && c <= 126).ToArray());
            return asciiString;
        }
        catch
        {
            return "<Invalid Hex Data>";
        }
    }

    // Organize tags into docks.
    public Dictionary<int, List<object>> GetTagsByDock()
    {
        var tagsByDock = new Dictionary<int, List<object>>
        {
            { 5, new List<object>() },
            { 6, new List<object>() },
            { 7, new List<object>() },
            { 8, new List<object>() }
        };

        foreach (var tag in _tags.Values)
        {
            string epcAscii = HexStringToAscii(tag.EPC_Hex);
            if (!(epcAscii.Contains("FF") || epcAscii.Contains("PL") ||
                  epcAscii.Contains("GL") || epcAscii.Contains("BL")))
            {
                continue;
            }

            var formattedTag = new
            {
                ScannedIndex = tag.ScannedIndex,
                EPC_Hex = tag.EPC_Hex,
                EPC_Ascii = epcAscii,
                tag.AntennaID,
                tag.PeakRSSI,
                Lot = tag.Lot ?? "N/A",
                Type = tag.Type ?? "N/A",
                Color = tag.Color ?? "N/A",
                Format = tag.Format ?? "N/A",
                Pounds = tag.Pounds ?? "N/A"
            };

            if (tag.Reader == "Primary")
            {
                if (tag.AntennaID >= 1 && tag.AntennaID <= 4)
                    tagsByDock[5].Add(formattedTag);
                else if (tag.AntennaID >= 5 && tag.AntennaID <= 8)
                    tagsByDock[6].Add(formattedTag);
            }
            else if (tag.Reader == "Third")
            {
                if (tag.AntennaID >= 1 && tag.AntennaID <= 4)
                    tagsByDock[7].Add(formattedTag);
                else if (tag.AntennaID >= 5 && tag.AntennaID <= 8)
                    tagsByDock[8].Add(formattedTag);
            }
        }

        return tagsByDock;
    }

    // Export tag data to Excel.
    public byte[] CreateExcelFileBytes(int dockDoor, string billOfLading)
    {
        var tagsByDock = GetTagsByDock();
        if (!tagsByDock.ContainsKey(dockDoor) || tagsByDock[dockDoor].Count == 0)
        {
            throw new Exception($"No tags available for Dock {dockDoor}");
        }

        var tags = tagsByDock[dockDoor];
        string directoryPath = @"C:\RFIDExports";
        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        string filePath = Path.Combine(directoryPath, $"Dock_{dockDoor}_{billOfLading}.xlsx");

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet($"Truck {dockDoor}");
            ws.Cell(1, 1).Value = "EPC ASCII";
            ws.Cell(1, 2).Value = "Scanned Index";

            int row = 2;
            foreach (var t in tags)
            {
                string epcAscii = t.GetType().GetProperty("EPC_Ascii")?.GetValue(t, null)?.ToString() ?? "<Unknown>";
                object? scannedIndexObj = t.GetType().GetProperty("ScannedIndex")?.GetValue(t, null);
                string scannedIndex = scannedIndexObj?.ToString() ?? "";
                ws.Cell(row, 1).Value = epcAscii;
                ws.Cell(row, 2).Value = scannedIndex;
                row++;
            }

            workbook.SaveAs(filePath);
            Console.WriteLine($"File saved: {filePath}");
        }

        if (!File.Exists(filePath))
        {
            throw new Exception($"File was not created: {filePath}");
        }

        return File.ReadAllBytes(filePath);
    }
}
