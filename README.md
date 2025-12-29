# Inventory Scanning System

This repository contains two cooperating services for managing RFID-based inventory scans:

1. **Node.js FXR90 gateway and web UI** (`app.js` + `Public/`) that:
   - Accepts push data from FXR90 readers at `/fxr90`.
   - Exposes start/stop/pause/clear controls for scanning.
   - Streams live tag updates to the browser via Socket.IO.
   - Exports scanned tags to Excel (raw or joined with a local DB workbook).
2. **.NET 6 RFID hub** (`Program.cs` and `RFIDService.cs`) that:
   - Connects directly to Zebra/FX series readers through the Symbol SDK.
   - Broadcasts tag reads over SignalR (`/rfidHub`) for dashboards and exports per dock door.
   - Enriches tag data from a cached Excel workbook (`DB.xlsx`).

Both services can run independently; use the Node gateway for the FXR90 REST workflow and browser UI, and use the .NET hub when you need the Symbol SDK integration and dock-door routing.

## Prerequisites

- **Node.js** 18+ (for the FXR90 gateway and browser UI).
- **.NET 6 SDK** (for the Symbol-based hub).
- **Hardware**: FXR90 readers reachable on your network, plus the Symbol SDK DLLs already provided in `libs/`.
- **Excel data**:
  - Node service: `data/Inventory.xlsx` (override with `LOCAL_DB_XLSX`).
  - .NET hub: `DB.xlsx` in the application directory.

## Configure the FXR90 Node gateway

1. Edit the reader definitions at the top of `app.js` (`READERS` array). Set the IPs and credentials for each FXR90.
2. (Optional) Point the data join to a different workbook by setting `LOCAL_DB_XLSX` before starting the server.
3. Ensure FXR90 readers are configured to push JSON tag events to `http://<server>:3000/fxr90`.

## Install and run the Node service

```bash
npm install
npm start        # launches the gateway on port 3000
```

Open the browser UI at `http://localhost:3000/`. Available pages in `Public/` include:
- `index.html`: Live scan dashboard (start/stop/pause/clear, export, joined export).
- `findBox.html`: “Find a Box” focused scanning with live tag updates.
- `dbPreview.html`: Preview the local Excel data the Node service will join against.

Key HTTP endpoints:
- `GET /api/start|stop|pause|clear`: Control scanning sessions.
- `GET /api/tags`: Current in-memory tags (no caching).
- `GET /api/export` and `GET /api/exportJoined`: Export raw or joined tag data to XLSX.
- `POST /fxr90`: Reader push endpoint for tag events (expects JSON array payloads).

## Install and run the .NET RFID hub

> The Symbol SDK native DLL (`RFIDAPI32PC.dll`) means the .NET service typically runs on Windows with reader drivers installed.

1. Ensure `DB.xlsx` is present alongside the executable (copied automatically on build).
2. Restore and build:
   ```bash
   dotnet restore
   dotnet build
   ```
3. Run the service:
   ```bash
   dotnet run
   ```
   The app listens on the URLs configured in `Program.cs` (defaults include `http://localhost:5026`).

Core APIs:
- `POST /api/reader/start` and `POST /api/reader/stop`: Begin or end reader inventory.
- `POST /api/reader/clear/{dockDoor}`: Clear tags for a specific dock (5–8).
- `GET /api/reader/tags`: Snapshot of tags grouped by dock door.
- `POST /api/reader/export/{dockDoor}/{billOfLading}`: Export dock-specific tags to Excel.
- `POST /api/reader/refreshExcelCache`: Force reload of `DB.xlsx` (expects `X-Refresh-Password` header).
- `GET /api/db/preview`: Diagnostics view of the cached Excel data.

## Logging and troubleshooting

- Node gateway logs incoming FXR90 payloads and scan state changes to stdout.
- If the Node export endpoints fail, verify the Excel source path and headers (joined exports require an `RFIDBOX` column).
- The .NET hub writes reader connection and tag broadcast messages to the console; ensure the configured IPs in `Program.cs` match your readers.
- Socket.IO (Node) and SignalR (.NET) clients both emit/receive real-time tag events—open the browser console to watch live traffic when debugging.
