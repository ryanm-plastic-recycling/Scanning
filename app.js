/****************************************************** 
 * app.js
 * 
 * Node.js server that:
 * 1) Hosts a push endpoint for FXR90 readers (/fxr90).
 * 2) Hosts a GUI at http://yourIP:3000/ (serves /public/index.html).
 * 3) Provides Start/Stop/Pause/Clear endpoints for the front-end.
 * 4) Merges and stores tag data in memory (unique tags only).
 * 5) Exports data to XLSX on demand.
 ******************************************************/

const express = require('express');
const axios = require('axios');
const https = require('https');
const http = require('http');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

// Create an HTTPS Agent that ignores self-signed certs (for testing only)
const httpsAgent = new https.Agent({
  rejectUnauthorized: false
});

// === CONFIGURE YOUR READERS HERE === //
const READERS = [
  {
    name: "Reader8Port",
    ip: "192.168.48.251",   //WIRED
    //ip: "192.168.49.24",   //WIFI
    user: "admin",
    pass: "PRIscan123!",
    token: null // Token for this reader will be stored here
  },
  {
    name: "Reader4Port",
    ip: "192.168.50.250",   //WIRED
    //ip: "192.168.50.117",   //WIFI
    user: "admin",
    pass: "PRIscan123!",
    token: null
  }
];

const APP_PORT = 3000;          // Node server port
const AUTH_URL = "/cloud/localRestLogin";
const START_URL = "/cloud/start";
const STOP_URL = "/cloud/stop";
const HEALTH_POLL_INTERVAL_MS = 7000; // 7s base interval
const MAX_HEALTH_BACKOFF_MS = 30000;  // cap backoff at 30s
const MIN_HEALTH_POLL_MS = 5000;      // do not poll faster than 5s

// Per-reader health tracking
const readerHealth = new Map();

function getReaderKey(reader = {}) {
  return reader.name || reader.ip || "unknown";
}

function ensureHealth(reader) {
  const key = getReaderKey(reader);
  if (!readerHealth.has(key)) {
    readerHealth.set(key, {
      name: reader.name,
      ip: reader.ip,
      reachable: false,
      authOk: false,
      isScanning: false,
      lastOkAt: null,
      lastError: "",
      lastTagAt: null,
      tokenIssuedAt: null,
      tokenAgeSec: null,
      failureCount: 0,
      nextPollTime: 0
    });
  }
  return readerHealth.get(key);
}

function healthPayload() {
  const now = Date.now();
  return Array.from(readerHealth.values()).map(h => ({
    name: h.name,
    ip: h.ip,
    reachable: !!h.reachable,
    authOk: !!h.authOk,
    isScanning: !!h.isScanning || !!isScanning || !!isFindBoxScanning,
    lastOkAt: h.lastOkAt,
    lastError: h.lastError,
    lastTagAt: h.lastTagAt,
    tokenIssuedAt: h.tokenIssuedAt,
    tokenAgeSec: h.tokenIssuedAt ? Math.floor((now - h.tokenIssuedAt) / 1000) : null
  }));
}

function emitHealth() {
  io.emit('readerHealth', healthPayload());
}

// Update helper that also emits
function updateHealth(reader, updates = {}) {
  const h = ensureHealth(reader);
  if (updates.isScanning === undefined) {
    updates.isScanning = isScanning || isFindBoxScanning;
  }
  Object.assign(h, updates);
  emitHealth();
  return h;
}

// Poll readers periodically to verify reachability/auth
async function pollHealth() {
  const now = Date.now();
  const tasks = READERS.map(async reader => {
    const h = ensureHealth(reader);
    if (h.nextPollTime && now < h.nextPollTime) return;
    try {
      await loginToReader(reader, { healthCheck: true });
      h.failureCount = 0;
      h.lastOkAt = Date.now();
      h.lastError = "";
      h.reachable = true;
      h.authOk = true;
      h.isScanning = isScanning || isFindBoxScanning;
      if (!h.tokenIssuedAt) h.tokenIssuedAt = h.lastOkAt;
      h.nextPollTime = now + Math.max(MIN_HEALTH_POLL_MS, HEALTH_POLL_INTERVAL_MS);
    } catch (err) {
      h.failureCount = (h.failureCount || 0) + 1;
      h.reachable = false;
      h.authOk = false;
      h.isScanning = isScanning || isFindBoxScanning;
      h.lastError = String(err?.message || err);
      const backoff = Math.min(MAX_HEALTH_BACKOFF_MS, HEALTH_POLL_INTERVAL_MS * (h.failureCount + 1));
      h.nextPollTime = now + backoff;
    }
  });
  await Promise.all(tasks);
  emitHealth();
}

// Local Excel source for preview (override with env var LOCAL_DB_XLSX)
const LOCAL_DB_XLSX = process.env.LOCAL_DB_XLSX || path.resolve(__dirname, "data", "Inventory.xlsx");

// In-memory store of unique tags (keyed by TID if available, else EPC)
let tagStore = {};

// Track scanning state & timer for normal scanning
let isScanning = false;
let scanStartTime = null;

// NEW: Flag for Find a Box scanning
let isFindBoxScanning = false;

const app = express();

app.post('/fxr90', (req, res, next) => {
  console.log("HEADERS:", req.headers);
  next(); // proceed to json parser
});

app.use('/fxr90', express.json({
  limit: '10mb',
  verify: (req, res, buf, encoding) => {
    if (!req.headers['content-type']?.includes('application/json')) {
      throw new Error('Invalid content-type');
    }
  }
}));

app.use(express.static('public')); // Serve static files from /public

// Create HTTP server and integrate Socket.IO for real-time updates
const server = http.createServer(app);
const { Server } = require('socket.io');
const io = new Server(server);

// Socket.IO connection handler
io.on('connection', (socket) => {
  console.log('A client connected. Socket ID:', socket.id);
  // Optionally, send the initial tag data:
  socket.emit('initialTags', Object.values(tagStore));
  socket.emit('readerHealth', healthPayload());

  // Listen for Find a Box scanning events from clients
  socket.on('startFindBoxScan', (data) => {
    const filter = data.filter || "";
    console.log("Received 'startFindBoxScan' event with filter:", filter);
    isFindBoxScanning = true;
    // For Find a Box scanning, force-start a new scanning session on readers
    Promise.all(READERS.map(r => loginAndStart(r)))
      .then(() => {
        // We do not set isScanning here since this is a dedicated find box session.
        scanStartTime = new Date();
        console.log("Find a Box scanning started successfully on all readers.");
        READERS.forEach(r => updateHealth(r, { isScanning: true }));
        socket.emit('findBoxScanStatus', { success: true, message: "Find a Box scanning started" });
      })
      .catch(err => {
        console.error("Error starting Find a Box scanning:", err);
        socket.emit('findBoxScanStatus', { success: false, message: err.toString() });
      });
  });

  socket.on('stopFindBoxScan', () => {
    console.log("Received 'stopFindBoxScan' event.");
    Promise.all(READERS.map(r => stopReader(r)))
      .then(() => {
        isFindBoxScanning = false;
        console.log("Find a Box scanning stopped successfully on all readers.");
        READERS.forEach(r => updateHealth(r, { isScanning: false }));
        socket.emit('findBoxScanStatus', { success: true, message: "Find a Box scanning stopped" });
      })
      .catch(err => {
        console.error("Error stopping Find a Box scanning:", err);
        socket.emit('findBoxScanStatus', { success: false, message: err.toString() });
      });
  });
});

// Initialize health records and start polling
READERS.forEach(r => ensureHealth(r));
setInterval(() => {
  pollHealth().catch(err => console.error("Health poll error:", err));
}, HEALTH_POLL_INTERVAL_MS);
pollHealth().catch(err => console.error("Initial health poll error:", err));

/******************************************************
 * 1) PUSH ENDPOINT for Readers to POST Tag Data
 ******************************************************/
app.post('/fxr90', (req, res) => {
  console.log("FXR90 data inbound:", JSON.stringify(req.body));

  const readerFromReq = resolveReaderFromRequest(req);
  if (readerFromReq) {
    updateHealth(readerFromReq, {
      lastTagAt: Date.now(),
      lastOkAt: Date.now(),
      reachable: true,
      authOk: true
    });
  } else {
    // Update a global entry so the UI still shows activity
    updateHealth({ name: "Unknown", ip: "unknown" }, { lastTagAt: Date.now() });
  }

  // Process tag data if either normal scanning or find box scanning is active.
  if (!isScanning && !isFindBoxScanning) {
    console.log("Received tag data but neither normal scanning nor find box scanning is active. Ignoring.");
    return res.sendStatus(200);
  }

  // Expect the incoming payload to be an array of events
  const events = req.body;
  if (!Array.isArray(events)) {
    return res.sendStatus(400);
  }

  events.forEach(event => {
    // Get the EPC hex from event.data.idHex
    const epcHex = (event?.data?.idHex || "").toUpperCase().trim();
    // If a separate TID exists, assign it; otherwise, leave it empty.
    const tidHex = "";
    const epcAscii = hexToAscii(epcHex);
    console.log(`Converted ${epcHex} to ASCII: ${epcAscii}`);

    // Optionally filter out tags that don't meet your criteria:
    if (!isValidAsciiTag(epcAscii)) return;

    // Use TID if available; otherwise, EPC as key
    const key = tidHex ? tidHex : epcHex;

    if (!tagStore[key]) {
      const newTag = {
        firstSeen: new Date().toISOString(),
        tidHex: tidHex,
        epcHex: epcHex,
        epcAscii: epcAscii,
        antenna: event.data.antenna,
        rssi: event.data.peakRssi,
        reader: readerFromReq?.name,
        readerIp: readerFromReq?.ip,
        seenCount: 1
      };
      tagStore[key] = newTag;
      // Emit events to connected clients via Socket.IO
      io.emit('newTag', newTag);
      io.emit('findBoxTag', newTag);
    } else {
      // Update the tag data and emit update for findBoxTag
      tagStore[key].seenCount = event.data.seenCount || (tagStore[key].seenCount + 1);
      tagStore[key].rssi = event.data.peakRssi;
      tagStore[key].reader = readerFromReq?.name || tagStore[key].reader;
      tagStore[key].readerIp = readerFromReq?.ip || tagStore[key].readerIp;
      io.emit('findBoxTag', tagStore[key]);
    }
  });
  res.sendStatus(200);
});

/******************************************************
 * 2) GUI CONTROL ENDPOINTS
 ******************************************************/

// Start scanning (normal scanning)
app.get('/api/start', async (req, res) => {
  console.log("Received /api/start request");
  try {
    await Promise.all(READERS.map(r => loginAndStart(r)));
    isScanning = true;
    READERS.forEach(r => updateHealth(r, { isScanning: true }));
    scanStartTime = new Date();
    console.log("Scanning started successfully on all readers.");
    res.json({ success: true, message: "Scanning started" });
  } catch (err) {
    console.error("Error in /api/start:", err);
    res.json({ success: false, message: err.toString() });
  }
});

// Stop scanning (normal scanning)
app.get('/api/stop', async (req, res) => {
  console.log("Received /api/stop request");
  try {
    await Promise.all(READERS.map(r => stopReader(r)));
    isScanning = false;
    READERS.forEach(r => updateHealth(r, { isScanning: false }));
    console.log("Scanning stopped successfully on all readers.");
    res.json({ success: true, message: "Scanning stopped" });
  } catch (err) {
    console.error("Error in /api/stop:", err);
    res.json({ success: false, message: err.toString() });
  }
});

// Pause scanning (data is retained)
app.get('/api/pause', async (req, res) => {
  console.log("Received /api/pause request");
  try {
    await Promise.all(READERS.map(r => stopReader(r)));
    isScanning = false;
    READERS.forEach(r => updateHealth(r, { isScanning: false }));
    res.json({ success: true, message: "Scanning paused (data retained)" });
  } catch (err) {
    console.error("Error in /api/pause:", err);
    res.json({ success: false, message: err.toString() });
  }
});

// Clear tag data
app.get('/api/clear', (req, res) => {
  tagStore = {};
  scanStartTime = isScanning || isFindBoxScanning ? new Date() : null;
  res.json({ success: true, message: "Data cleared" });
});

// Get all current tag data with caching disabled
app.get('/api/tags', (req, res) => {
  res.set('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.json(Object.values(tagStore));
});

// Reader health
app.get('/api/health', (req, res) => {
  res.set('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.json(healthPayload());
});

// Export to XLSX
app.get('/api/export', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('FXR90 Tags');
    sheet.columns = [
      { header: 'First Seen', key: 'firstSeen', width: 20 },
      { header: 'TID (Hex)', key: 'tidHex', width: 35 },
      { header: 'EPC (Hex)', key: 'epcHex', width: 35 },
      { header: 'EPC (ASCII)', key: 'epcAscii', width: 35 },
      { header: 'RSSI', key: 'rssi', width: 10 },
      { header: 'Seen Count', key: 'seenCount', width: 10 }
    ];
    Object.values(tagStore).forEach(t => {
      sheet.addRow(t);
    });
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="fxr90-tags.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err) {
    console.error("Error exporting XLSX:", err);
    res.status(500).send(err.toString());
  }
});

// Export to XLSX (JOINED with DB columns)
app.get('/api/exportJoined', async (req, res) => {
  try {
    // 1) Build DB lookup keyed by RFIDBOX (ASCII)
    if (!fs.existsSync(LOCAL_DB_XLSX)) {
      return res.status(500).send(`Local Excel file not found: ${LOCAL_DB_XLSX}`);
    }

    const dbWb = new ExcelJS.Workbook();
    await dbWb.xlsx.readFile(LOCAL_DB_XLSX);
    const ws = dbWb.worksheets[0];
    if (!ws) return res.status(500).send("No worksheet found in local Excel file.");

    // Read headers
    const headerRow = ws.getRow(1).values.slice(1).map(v => String(v ?? "").trim());
    const headerIndex = new Map();
    headerRow.forEach((h, idx) => headerIndex.set(String(h || "").trim().toUpperCase(), idx));

    // RFIDBOX is the ASCII key column (you confirmed)
    const keyColIdx = headerIndex.get("RFIDBOX");
    if (keyColIdx === undefined) {
      return res.status(500).send("DB sheet missing required header column 'RFIDBOX'.");
    }

    // Helper to safely read a cell by header name
    const getByHeader = (rowArr, name) => {
      const idx = headerIndex.get(String(name).trim().toUpperCase());
      if (idx === undefined) return "";
      const v = rowArr[idx];
      return (v === null || v === undefined) ? "" : String(v);
    };

    const dbMap = new Map(); // ASCII -> object
    for (let r = 2; r <= ws.rowCount; r++) {
      const rowArr = ws.getRow(r).values.slice(1);
      if (rowArr.every(v => v === null || v === undefined || String(v).trim() === "")) continue;

      const key = (rowArr[keyColIdx] ?? "").toString().trim().toUpperCase();
      if (!key) continue;

      // Store full row so we can pull any columns you want by header
      dbMap.set(key, {
        lot:      getByHeader(rowArr, "LOT") || getByHeader(rowArr, "RFID") || "",
        dept:     getByHeader(rowArr, "DEPT") || getByHeader(rowArr, "DEPARTMENT") || "",
        row:      getByHeader(rowArr, "ROW") || "",
        deptLot:  getByHeader(rowArr, "DEPT LOT") || getByHeader(rowArr, "DEPTLOT") || "",
        supplier: getByHeader(rowArr, "SUPPLIER") || getByHeader(rowArr, "CUSTOMER") || "",
        type:     getByHeader(rowArr, "TYPE") || getByHeader(rowArr, "MATERIAL") || "",
        color:    getByHeader(rowArr, "COLOR") || "",
        format:   getByHeader(rowArr, "FORMAT") || "",
        pounds:   getByHeader(rowArr, "POUNDS") || getByHeader(rowArr, "WEIGHT") || "",
        price:    getByHeader(rowArr, "PRICE") || "",
        freight:  getByHeader(rowArr, "FREIGHT") || "",
        toll:     getByHeader(rowArr, "TOLL") || getByHeader(rowArr, "TOLLING") || "",
        date:     getByHeader(rowArr, "DATE") || ""
      });
    }

    // 2) Create workbook for export
    const outWb = new ExcelJS.Workbook();
    const sheet = outWb.addWorksheet('FXR90 Tags (Joined)');

    sheet.columns = [
      { header: 'First Seen', key: 'firstSeen', width: 22 },
      { header: 'TID (Hex)', key: 'tidHex', width: 35 },
      { header: 'EPC (Hex)', key: 'epcHex', width: 35 },
      { header: 'EPC (ASCII)', key: 'epcAscii', width: 20 },
      { header: 'RSSI', key: 'rssi', width: 10 },
      { header: 'Seen Count', key: 'seenCount', width: 12 },

      { header: 'LOT', key: 'lot', width: 16 },
      { header: 'DEPT', key: 'dept', width: 12 },
      { header: 'ROW', key: 'row', width: 12 },
      { header: 'DEPT LOT', key: 'deptLot', width: 18 },
      { header: 'SUPPLIER', key: 'supplier', width: 22 },
      { header: 'TYPE', key: 'type', width: 18 },
      { header: 'COLOR', key: 'color', width: 14 },
      { header: 'FORMAT', key: 'format', width: 14 },
      { header: 'POUNDS', key: 'pounds', width: 12 },
      { header: 'PRICE', key: 'price', width: 12 },
      { header: 'FREIGHT', key: 'freight', width: 12 },
      { header: 'TOLL', key: 'toll', width: 12 },
      { header: 'DATE', key: 'date', width: 16 }
    ];

    // 3) Add rows by joining tag.epcAscii -> dbMap(RFIDBOX)
    const tags = Object.values(tagStore);
    for (const t of tags) {
      const asciiKey = (t.epcAscii || "").toString().trim().toUpperCase();
      const db = dbMap.get(asciiKey) || {};

      sheet.addRow({
        firstSeen: t.firstSeen || "",
        tidHex: t.tidHex || "",
        epcHex: t.epcHex || "",
        epcAscii: t.epcAscii || "",
        rssi: t.rssi ?? "",
        seenCount: t.seenCount ?? "",

        lot: db.lot || "",
        dept: db.dept || "",
        row: db.row || "",
        deptLot: db.deptLot || "",
        supplier: db.supplier || "",
        type: db.type || "",
        color: db.color || "",
        format: db.format || "",
        pounds: db.pounds || "",
        price: db.price || "",
        freight: db.freight || "",
        toll: db.toll || "",
        date: db.date || ""
      });
    }

    const buffer = await outWb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="fxr90-tags-joined.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err) {
    console.error("Error exporting JOINED XLSX:", err);
    res.status(500).send(String(err?.message || err));
  }
});

/******************************************************
 * DB PREVIEW ENDPOINT (Node on 3000)
 * Reads a LOCAL Excel file and returns { columns, rows }
 ******************************************************/
app.get("/api/db/preview", async (req, res) => {
  try {
    if (!fs.existsSync(LOCAL_DB_XLSX)) {
      return res.status(500).json({
        error: `Local Excel file not found: ${LOCAL_DB_XLSX}. Set env var LOCAL_DB_XLSX to override.`
      });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_DB_XLSX);

    const ws = workbook.worksheets[0];
    if (!ws) return res.status(500).json({ error: "No worksheet found in file." });

    const headerRow = ws.getRow(1);
    const columns = headerRow.values.slice(1).map(v => String(v ?? "").trim());

    const rows = [];
    const maxRows = 200;

    for (let r = 2; r <= ws.rowCount && rows.length < maxRows; r++) {
      const row = ws.getRow(r).values.slice(1);
      if (row.every(v => v === null || v === undefined || String(v).trim() === "")) continue;
      rows.push(row);
    }

    res.json({ columns, rows, file: LOCAL_DB_XLSX });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

/******************************************************
 * DB FULL ENDPOINT (Node on 3000)
 * Reads the FULL local Excel file and returns { columns, rows }
 * Use this for index.html joins (not for rendering huge tables).
 ******************************************************/
app.get("/api/db/all", async (req, res) => {
  try {
    if (!fs.existsSync(LOCAL_DB_XLSX)) {
      return res.status(500).json({
        error: `Local Excel file not found: ${LOCAL_DB_XLSX}. Set env var LOCAL_DB_XLSX to override.`
      });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(LOCAL_DB_XLSX);

    const ws = workbook.worksheets[0];
    if (!ws) return res.status(500).json({ error: "No worksheet found in file." });

    const headerRow = ws.getRow(1);
    const columns = headerRow.values.slice(1).map(v => String(v ?? "").trim());

    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r).values.slice(1);
      // skip fully empty rows
      if (row.every(v => v === null || v === undefined || String(v).trim() === "")) continue;
      rows.push(row);
    }

    res.json({ columns, rows, file: LOCAL_DB_XLSX, totalRows: rows.length });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

/******************************************************
 * HELPER FUNCTIONS
 ******************************************************/

// Log in to the reader, then force-stop any active scan before starting a new scan
async function loginAndStart(reader) {
  const loginUrl = `https://${reader.ip}${AUTH_URL}`;
  const authString = Buffer.from(`${reader.user}:${reader.pass}`).toString('base64');
  const resp = await axios.get(loginUrl, {
    headers: { 'Authorization': `Basic ${authString}` },
    httpsAgent
  });
  if (resp.data && resp.data.message) {
    reader.token = resp.data.message;
    const now = Date.now();
    updateHealth(reader, { reachable: true, authOk: true, lastOkAt: now, lastError: "", tokenIssuedAt: now });
  } else {
    throw new Error(`No token in login response from ${reader.ip}`);
  }
  // Force-stop any active scan before starting
  try {
    await stopReader(reader);
  } catch (err) {
    console.log(`Stop command not successful (possibly no active scan) for ${reader.ip}: ${err}`);
  }
  // Wait a few seconds to let the reader settle
  await new Promise(resolve => setTimeout(resolve, 3000));
  await startReader(reader);
}

// Helper for re-login without starting scanning (used if token is rejected during stop)
async function loginToReader(reader, opts = {}) {
  const { healthCheck = false } = opts;
  const loginUrl = `https://${reader.ip}${AUTH_URL}`;
  const authString = Buffer.from(`${reader.user}:${reader.pass}`).toString('base64');
  const h = ensureHealth(reader);
  const resp = await axios.get(loginUrl, {
    headers: { 'Authorization': `Basic ${authString}` },
    httpsAgent
  });
  if (resp.data && resp.data.message) {
    if (!healthCheck || !reader.token) {
      reader.token = resp.data.message;
    }
    const now = Date.now();
    const updates = {
      reachable: true,
      authOk: true,
      lastOkAt: now,
      lastError: ""
    };
    if (!healthCheck || !h.tokenIssuedAt) {
      updates.tokenIssuedAt = now;
    }
    updateHealth(reader, updates);
  } else {
    throw new Error(`No token in login response from ${reader.ip}`);
  }
}

// Call /cloud/start; if a 422 error occurs (indicating scan already running), stop then retry
async function startReader(reader) {
  const url = `https://${reader.ip}${START_URL}`;
  try {
    await axios.put(url, {}, {
      headers: { 'Authorization': `Bearer ${reader.token}` },
      httpsAgent
    });
    updateHealth(reader, { isScanning: true, reachable: true, authOk: true, lastOkAt: Date.now(), lastError: "" });
  } catch (err) {
    if (err.response && err.response.status === 422) {
      console.log(`Reader ${reader.ip} already scanning. Issuing STOP then retrying START...`);
      await stopReader(reader);
      await new Promise(resolve => setTimeout(resolve, 3000)); // extra delay
      await axios.put(url, {}, {
        headers: { 'Authorization': `Bearer ${reader.token}` },
        httpsAgent
      });
      updateHealth(reader, { isScanning: true, reachable: true, authOk: true, lastOkAt: Date.now(), lastError: "" });
    } else {
      throw err;
    }
  }
}

// Call /cloud/stop; if a 401 occurs, re-login and retry
async function stopReader(reader) {
  const url = `https://${reader.ip}${STOP_URL}`;
  try {
    const resp = await axios.put(url, {}, {
      headers: { 'Authorization': `Bearer ${reader.token}` },
      httpsAgent
    });
    updateHealth(reader, { isScanning: false, reachable: true, authOk: true, lastOkAt: Date.now(), lastError: "" });
    return resp;
  } catch (err) {
    if (err.response && err.response.status === 401) {
      console.log(`Received 401 on stop for ${reader.ip}, re-logging in...`);
      await loginToReader(reader);
      const resp = await axios.put(url, {}, {
        headers: { 'Authorization': `Bearer ${reader.token}` },
        httpsAgent
      });
      updateHealth(reader, { isScanning: false, reachable: true, authOk: true, lastOkAt: Date.now(), lastError: "" });
      return resp;
    }
    throw err;
  }
}

function resolveReaderFromRequest(req) {
  const forwarded = (req.headers['x-forwarded-for'] || "").split(',')[0].trim();
  const rawRemote = req.connection?.remoteAddress || req.socket?.remoteAddress || "";
  const cleanedRemote = rawRemote.replace("::ffff:", "");
  const ipGuess = forwarded || cleanedRemote;
  return READERS.find(r => ipGuess.endsWith(r.ip));
}

// Convert EPC from hex to ASCII
function hexToAscii(hexString) {
  hexString = hexString.replace(/\s+/g, '');
  if (hexString.length % 2 !== 0) return "";
  let ascii = "";
  for (let i = 0; i < hexString.length; i += 2) {
    const part = hexString.substr(i, 2);
    const code = parseInt(part, 16);
    ascii += (code >= 32 && code <= 126) ? String.fromCharCode(code) : ".";
  }
  return ascii;
}

// Check if ASCII starts with "FF", "PL", "BL", "GL", or "PRI-"
function isValidAsciiTag(asciiTag) {
  if (!asciiTag) return false;
  const upper = asciiTag.toUpperCase();
  return (
    upper.startsWith("FF") ||
    upper.startsWith("PL") ||
    upper.startsWith("BL") ||
    upper.startsWith("GL") ||
    upper.startsWith("PRI-")
  );
}

app.use((err, req, res, next) => {
  if (err.type === 'entity.parse.failed') {
    console.error('Invalid JSON payload received.');
    return res.status(400).send('Invalid JSON');
  }

  if (err.type === 'request.aborted') {
    console.error('Request was aborted by the client.');
    return res.status(400).send('Request aborted');
  }

  next(err); // pass other errors along
});

/******************************************************
 * 3) START the EXPRESS SERVER
 ******************************************************/
server.listen(APP_PORT, () => {
  console.log(`Server listening on port ${APP_PORT}`);
  console.log(`Point your browser to http://localhost:${APP_PORT}/`);
  console.log(`DB Preview reads: ${LOCAL_DB_XLSX} (override with LOCAL_DB_XLSX env var)`);
});
