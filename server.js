const express = require('express');
const http = require('http');
const WebSocket = require('ws');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');
const fetch = require('node-fetch');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const LOGS_DIR = path.join(__dirname, 'logs');
if (!fs.existsSync(LOGS_DIR)) fs.mkdirSync(LOGS_DIR);
const OPERATORS_PATH = path.join(__dirname, 'data', 'operators.json');
function loadOperators() {
  try { return JSON.parse(fs.readFileSync(OPERATORS_PATH,'utf-8')); } catch(e) { return []; }
}
function saveOperators(list) {
  fs.writeFileSync(OPERATORS_PATH, JSON.stringify([...new Set(list)].filter(Boolean), null, 2));
}

// ── Config (persisted to config.json) ────────────────────────────────────────
const CONFIG_PATH = path.join(__dirname, 'config.json');
let config = {
  vmixHost: '127.0.0.1',
  vmixPort: 8088,
  inputName: 'Lower Third',
  takeLiveDelay: 500,
  overlayChannel: 1,
  autoOverlay: false,
  vmixEnabled: true,
  // Mix routing: array of mix numbers (0 = Master, 1-16 = Mix 1-16)
  targetMixes: [0],
  mixPresets: [],
  activePreset: null,
  scPresets: [],
  activeScPreset: null,
  // Keyboard shortcuts: action -> key string
  shortcuts: {
    sendDataOnly:   'D',
    sendAndGoLive:  'G',
    clearAndHide:   'C',
    nextName:       'ArrowRight',
    prevName:       'ArrowLeft',
    overlayIn:      'I',
    overlayOut:     'O',
    takePreview:    'Enter'
  }
};
function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) {
      const saved = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));
      config = { ...config, ...saved };
    }
  } catch(e) {}
}
function saveConfig() {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
}
loadConfig();

// ── Names DB ──────────────────────────────────────────────────────────────────
let namesDb = [];
function loadNames() {
  try {
    const csvPath = path.join(__dirname, 'data', 'names.csv');
    const content = fs.readFileSync(csvPath, 'utf-8');
    namesDb = parse(content, { columns: true, skip_empty_lines: true });
    console.log(`Loaded ${namesDb.length} names`);
  } catch(e) { console.error('CSV load error:', e.message); }
}
function watchNames() {
  const csvPath = path.join(__dirname, 'data', 'names.csv');
  fs.watchFile(csvPath, { interval: 2000 }, () => {
    loadNames();
    broadcast({ type: 'db_reloaded', count: namesDb.length });
  });
}
loadNames();
watchNames();

// ── State ─────────────────────────────────────────────────────────────────────
let state = {
  live: { id: null, name: '', title: '', organisation: '', isBlank: true },
  preview: { id: null, name: '', title: '', organisation: '', isBlank: true },
  overlayLive: false,   // true = graphic is actually showing in VMix
  mode: 'remote',
  currentIndex: -1,
  history: [],
  sessionLog: [],
  sessionStart: new Date().toISOString(),
  vmixConnected: false
};

// ── VMix API ──────────────────────────────────────────────────────────────────
function vmixUrl(fn) {
  return `http://${config.vmixHost}:${config.vmixPort}/api/?Function=${fn}`;
}

async function vmixCall(fn) {
  try {
    const res = await fetch(vmixUrl(fn), { timeout: 3000 });
    const wasConnected = state.vmixConnected;
    state.vmixConnected = res.ok;
    if (state.vmixConnected !== wasConnected) broadcast({ type: 'vmix_status', connected: state.vmixConnected });
    return res.ok;
  } catch(e) {
    const wasConnected = state.vmixConnected;
    state.vmixConnected = false;
    if (wasConnected) broadcast({ type: 'vmix_status', connected: false, error: e.message });
    return false;
  }
}

// Fire overlay command on all targeted mixes
async function vmixOverlayIn() {
  const mixes = config.targetMixes && config.targetMixes.length ? config.targetMixes : [0];
  const ch = config.overlayChannel || 1;
  const input = encodeURIComponent(config.inputName);
  const results = await Promise.all(mixes.map(mix => {
    // mix 0 = Master (no mix suffix), 1-16 = Mix1..Mix16
    const fn = mix === 0
      ? `OverlayInput${ch}In&Input=${input}`
      : `OverlayInput${ch}In&Input=${input}&Mix=${mix}`;
    return vmixCall(fn);
  }));
  return results.every(Boolean);
}

async function vmixOverlayOut() {
  const mixes = config.targetMixes && config.targetMixes.length ? config.targetMixes : [0];
  const ch = config.overlayChannel || 1;
  const input = encodeURIComponent(config.inputName);
  const results = await Promise.all(mixes.map(mix => {
    const fn = mix === 0
      ? `OverlayInput${ch}Out&Input=${input}`
      : `OverlayInput${ch}Out&Input=${input}&Mix=${mix}`;
    return vmixCall(fn);
  }));
  return results.every(Boolean);
}

async function checkVmixConnection() {
  const prev = state.vmixConnected;
  try {
    const res = await fetch(`http://${config.vmixHost}:${config.vmixPort}/api/`, { timeout: 2000 });
    state.vmixConnected = res.ok;
  } catch(e) {
    state.vmixConnected = false;
  }
  if (state.vmixConnected !== prev) {
    broadcast({ type: 'vmix_status', connected: state.vmixConnected });
  }
}
setInterval(checkVmixConnection, 5000);
checkVmixConnection();

// ── VMix overlay status polling ───────────────────────────────────────────────
// Polls VMix XML every 3s to accurately check if our input is in an active overlay
// VMix XML structure: <overlay number="1"><input shortTitle="Lower Third" .../></overlay>
// An empty overlay looks like: <overlay number="1" /> or <overlay number="1"></overlay>
// Track last time we manually triggered overlay so poll doesn't fight us
async function pollOverlayStatus() {
  if (!config.vmixEnabled || !state.vmixConnected) return;
  try {
    const res = await fetch(`http://${config.vmixHost}:${config.vmixPort}/api/`, { timeout: 2000 });
    if (!res.ok) return;
    const xml = await res.text();
    const inputName = config.inputName.toLowerCase();
    const ch = config.overlayChannel || 1;

    // Log full XML once so user can paste it to help us tune the parser
    if (!state._xmlLogged) {
      state._xmlLogged = true;
      const overlayIdx = xml.toLowerCase().indexOf('overlay');
      if (overlayIdx >= 0) {
        console.log('[VMix XML - overlay section - paste this in chat]');
        console.log(xml.substring(Math.max(0, overlayIdx - 50), overlayIdx + 500));
        console.log('[end of sample]');
      }
    }

    // VMix overlay XML:
    //   Active:  <overlay number="1">1</overlay>  (content = input number)
    //   Empty:   <overlay number="1" />            (self-closing)
    // We just check if the overlay channel has ANY content — if it does, something is live.
    // We trust that if WE put something live, it's our graphic.

    const overlayPattern = '<overlay[^>]+number=["\']+' + ch + '["\']+[^>]*>\\s*(\\d+)\\s*<\\/overlay>';
    const overlayRe = new RegExp(overlayPattern, 'i');
    const overlayMatch = overlayRe.exec(xml);
    const isLive = overlayMatch !== null;

    if (isLive !== state.overlayLive) {
      console.log(`[Overlay poll] overlay${ch}: ${state.overlayLive} -> ${isLive}`);
      state.overlayLive = isLive;
      broadcast({ type: 'overlay_state', overlayLive: state.overlayLive });
    }
  } catch(e) {
    // Silent fail — polling is non-critical
  }
}
setInterval(pollOverlayStatus, 2000);

// ── Session logging ───────────────────────────────────────────────────────────
function logEntry(person, event, triggeredBy) {
  const entry = {
    id: person?.id || '',
    name: person?.name || '',
    title: person?.title || '',
    organisation: person?.organisation || '',
    event,
    triggeredBy,
    timestamp: new Date().toISOString(),
    time: new Date().toLocaleTimeString()
  };
  state.sessionLog.push(entry);
  return entry;
}

async function exportSessionExcel() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'VMix Title Controller';
  wb.created = new Date();

  const ws = wb.addWorksheet('Session Log');

  // Header row styling
  ws.columns = [
    { header: 'ID',           key: 'id',           width: 8  },
    { header: 'Name',         key: 'name',         width: 28 },
    { header: 'Title / Role', key: 'title',        width: 28 },
    { header: 'Organisation', key: 'organisation', width: 28 },
    { header: 'Event',        key: 'event',        width: 14 },
    { header: 'Time',         key: 'time',         width: 12 },
    { header: 'Timestamp',    key: 'timestamp',    width: 26 },
    { header: 'Triggered By', key: 'triggeredBy',  width: 16 },
  ];

  // Style header
  ws.getRow(1).eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2A6E' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF7C6AF5' } } };
  });
  ws.getRow(1).height = 22;

  // Data rows
  state.sessionLog.forEach((entry, i) => {
    const row = ws.addRow(entry);
    row.height = 18;
    const fill = i % 2 === 0
      ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F4FF' } }
      : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    row.eachCell(cell => {
      cell.fill = fill;
      cell.font = { name: 'Arial', size: 10 };
      cell.alignment = { vertical: 'middle' };
    });
    // Colour event cell
    const eventCell = row.getCell('event');
    if (entry.event === 'LIVE') {
      eventCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F6E56' } };
      eventCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE1F5EE' } };
    } else if (entry.event === 'CLEARED') {
      eventCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFA32D2D' } };
      eventCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCEBEB' } };
    } else if (entry.event === 'PREVIEW') {
      eventCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF185FA5' } };
      eventCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F1FB' } };
    }
  });

  // Summary sheet
  const ws2 = wb.addWorksheet('Summary');
  ws2.columns = [
    { header: 'Metric', key: 'metric', width: 30 },
    { header: 'Value',  key: 'value',  width: 40 },
  ];
  ws2.getRow(1).eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2A6E' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  ws2.getRow(1).height = 22;

  const liveEvents = state.sessionLog.filter(e => e.event === 'LIVE');
  const uniqueSpeakers = [...new Set(liveEvents.map(e => e.name).filter(Boolean))];

  [
    { metric: 'Session started', value: new Date(state.sessionStart).toLocaleString() },
    { metric: 'Session exported', value: new Date().toLocaleString() },
    { metric: 'Total log entries', value: state.sessionLog.length },
    { metric: 'Total times taken live', value: liveEvents.length },
    { metric: 'Unique speakers', value: uniqueSpeakers.length },
    { metric: 'Speakers', value: uniqueSpeakers.join(', ') },
  ].forEach((r, i) => {
    const row = ws2.addRow(r);
    row.height = 18;
    row.eachCell(cell => {
      cell.font = { name: 'Arial', size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i % 2 === 0 ? 'FFF5F4FF' : 'FFFFFFFF' } };
    });
  });

  const dateStr = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const filename = `session-${dateStr}.xlsx`;
  const filepath = path.join(LOGS_DIR, filename);
  await wb.xlsx.writeFile(filepath);
  return { filename, filepath };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function lookupById(id) {
  return namesDb.find(r => String(r.id) === String(id)) || null;
}

function personObj(person) {
  return person
    ? { id: person.id, name: person.name, title: person.title, organisation: person.organisation, isBlank: false }
    : { id: null, name: '', title: '', organisation: '', isBlank: true };
}

function broadcast(msg) {
  const data = JSON.stringify(msg);
  wss.clients.forEach(ws => { if (ws.readyState === WebSocket.OPEN) ws.send(data); });
}

function setLive(person, triggeredBy = 'operator') {
  // Deduplicate — don't broadcast if same person is already live
  const newId = person ? String(person.id) : null;
  const curId = state.live.isBlank ? null : String(state.live.id);
  if (newId === curId) return state.live; // no change, skip broadcast

  state.live = personObj(person);
  state.lastUpdated = new Date().toISOString();
  if (person) {
    state.currentIndex = namesDb.findIndex(r => String(r.id) === String(person.id));
    state.history.unshift({ ...state.live, time: state.lastUpdated });
    if (state.history.length > 30) state.history.pop();
    logEntry(person, 'LIVE', triggeredBy);
  } else {
    logEntry(null, 'CLEARED', triggeredBy);
  }
  broadcast({ type: 'title_changed', live: state.live, preview: state.preview, updatedBy: triggeredBy, timestamp: state.lastUpdated });
  return state.live;
}

function setPreview(person, triggeredBy = 'operator') {
  const newId = person ? String(person.id) : null;
  const curId = state.preview.isBlank ? null : String(state.preview.id);
  if (newId === curId) return state.preview; // no change

  state.preview = personObj(person);
  logEntry(person, 'PREVIEW', triggeredBy);
  broadcast({ type: 'preview_changed', preview: state.preview });
  return state.preview;
}

// ── REST API ──────────────────────────────────────────────────────────────────

// VMix text endpoints
app.get('/vmix/name',         (req, res) => res.type('text').send(state.live.name || ''));
app.get('/vmix/title',        (req, res) => res.type('text').send(state.live.title || ''));
app.get('/vmix/organisation', (req, res) => res.type('text').send(state.live.organisation || ''));
app.get('/vmix/isblank',      (req, res) => res.type('text').send(state.live.isBlank ? '1' : '0'));

app.get('/current',  (req, res) => res.json(state.live));
app.get('/db',       (req, res) => res.json(namesDb));
app.get('/history',  (req, res) => res.json(state.history));
app.get('/status',   (req, res) => res.json({
  ok: true, namesCount: namesDb.length, mode: state.mode,
  live: state.live, preview: state.preview,
  vmixConnected: state.vmixConnected, uptime: process.uptime(),
  sessionLog: state.sessionLog.length
}));

app.get('/lookup/:id', (req, res) => {
  const person = lookupById(req.params.id);
  res.json(person ? { found: true, person } : { found: false });
});

// Set preview
app.post('/preview/:id', (req, res) => {
  const person = lookupById(req.params.id);
  if (!person) return res.status(404).json({ error: 'ID not found' });
  const result = setPreview(person, req.body.operator || 'operator');
  res.json({ ok: true, preview: result });
});

// Set live (data only, no VMix trigger)
app.post('/set/:id', (req, res) => {
  const person = lookupById(req.params.id);
  if (!person) return res.status(404).json({ error: 'ID not found' });
  const result = setLive(person, req.body.operator || 'operator');
  res.json({ ok: true, live: result });
});

// Set live + trigger VMix overlay after delay
app.post('/take/:id', async (req, res) => {
  const person = lookupById(req.params.id);
  if (!person) return res.status(404).json({ error: 'ID not found' });
  setLive(person, req.body.operator || 'operator');
  res.json({ ok: true, live: state.live, vmixPending: true });
  if (config.vmixEnabled !== false) {
    setTimeout(async () => {
      const ok = await vmixOverlayIn();
      if (ok) { state.overlayLive = true; broadcast({ type: 'overlay_state', overlayLive: true }); }
    }, config.takeLiveDelay);
  }
});

// Take preview to live
app.post('/take-preview', async (req, res) => {
  if (state.preview.isBlank) return res.status(400).json({ error: 'Nothing in preview' });
  const person = lookupById(state.preview.id);
  if (!person) return res.status(400).json({ error: 'Preview person not found in DB' });
  // Force update even if same person — overlay may be off, data same
  state.live = personObj(person);
  state.lastUpdated = new Date().toISOString();
  state.currentIndex = namesDb.findIndex(r => String(r.id) === String(person.id));
  logEntry(person, 'LIVE', req.body.operator || 'operator');
  broadcast({ type: 'title_changed', live: state.live, preview: state.preview, updatedBy: req.body.operator || 'operator', timestamp: state.lastUpdated });
  state.preview = personObj(null);
  broadcast({ type: 'preview_changed', preview: state.preview });
  res.json({ ok: true, live: state.live });
  if (config.vmixEnabled !== false) {
    setTimeout(async () => {
      const ok = await vmixOverlayIn();
      if (ok) { state.overlayLive = true; broadcast({ type: 'overlay_state', overlayLive: true }); }
    }, config.takeLiveDelay);
  }
});

// VMix overlay in (graphic on screen)
app.post('/vmix/in', async (req, res) => {
  const ok = await vmixOverlayIn();
  if (ok) {
    state.overlayLive = true;
    broadcast({ type: 'overlay_state', overlayLive: true });
  }
  res.json({ ok, overlayLive: state.overlayLive, connected: state.vmixConnected });
});

// VMix overlay out (graphic off screen)
app.post('/vmix/out', async (req, res) => {
  const ok = await vmixOverlayOut();
  if (ok) {
    state.overlayLive = false;
    broadcast({ type: 'overlay_state', overlayLive: false });
    setLive(null, 'vmix-out');
  }
  res.json({ ok, overlayLive: state.overlayLive, connected: state.vmixConnected });
});

// Clear data + remove from screen
app.post('/clear', async (req, res) => {
  const withVmix = req.body.vmix !== false;
  if (withVmix) await vmixOverlayOut();
  state.overlayLive = false;
  broadcast({ type: 'overlay_state', overlayLive: false });
  const result = setLive(null, req.body.operator || 'operator');
  res.json({ ok: true, live: result });
});

// StreamDeck
app.post('/next', async (req, res) => {
  if (!namesDb.length) return res.status(400).json({ error: 'No names' });
  const idx = state.currentIndex < namesDb.length - 1 ? state.currentIndex + 1 : 0;
  setLive(namesDb[idx], 'streamdeck');
  res.json({ ok: true, live: state.live });
  setTimeout(() => vmixOverlayIn(), config.takeLiveDelay);
});
app.post('/prev', async (req, res) => {
  if (!namesDb.length) return res.status(400).json({ error: 'No names' });
  const idx = state.currentIndex > 0 ? state.currentIndex - 1 : namesDb.length - 1;
  setLive(namesDb[idx], 'streamdeck');
  res.json({ ok: true, live: state.live });
  setTimeout(() => vmixOverlayIn(), config.takeLiveDelay);
});

// Config
app.get('/config', (req, res) => res.json(config));
app.post('/config', (req, res) => {
  config = { ...config, ...req.body };
  saveConfig();
  broadcast({ type: 'config_changed', config });
  res.json({ ok: true, config });
});

// Export session log
app.post('/export', async (req, res) => {
  try {
    const { filename, filepath } = await exportSessionExcel();
    res.json({ ok: true, filename, path: filepath });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Download log file
app.get('/logs', (req, res) => {
  const files = fs.readdirSync(LOGS_DIR).filter(f => f.endsWith('.xlsx')).sort().reverse();
  res.json(files);
});
app.get('/logs/:filename', (req, res) => {
  const fp = path.join(LOGS_DIR, req.params.filename);
  if (!fs.existsSync(fp)) return res.status(404).send('Not found');
  res.download(fp);
});

// New session (reset log)
app.post('/new-session', (req, res) => {
  state.sessionLog = [];
  state.sessionStart = new Date().toISOString();
  broadcast({ type: 'session_reset', sessionStart: state.sessionStart });
  res.json({ ok: true });
});

// Save mix preset
app.post('/presets', (req, res) => {
  const { name, mixes } = req.body;
  if (!name || !Array.isArray(mixes)) return res.status(400).json({ error: 'name and mixes required' });
  config.mixPresets = config.mixPresets.filter(p => p.name !== name);
  config.mixPresets.push({ name, mixes });
  config.activePreset = name;
  saveConfig();
  broadcast({ type: 'config_changed', config });
  res.json({ ok: true, presets: config.mixPresets });
});

app.delete('/presets/:name', (req, res) => {
  config.mixPresets = config.mixPresets.filter(p => p.name !== decodeURIComponent(req.params.name));
  if (config.activePreset === decodeURIComponent(req.params.name)) config.activePreset = null;
  saveConfig();
  broadcast({ type: 'config_changed', config });
  res.json({ ok: true });
});

app.post('/presets/:name/apply', (req, res) => {
  const preset = config.mixPresets.find(p => p.name === decodeURIComponent(req.params.name));
  if (!preset) return res.status(404).json({ error: 'Preset not found' });
  config.targetMixes = preset.mixes;
  config.activePreset = preset.name;
  saveConfig();
  broadcast({ type: 'config_changed', config });
  res.json({ ok: true, config });
});

// Operators
app.get('/operators', (req, res) => res.json(loadOperators()));
app.post('/operators', (req, res) => {
  const { name } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: 'Name required' });
  const list = loadOperators();
  const trimmed = name.trim();
  if (!list.includes(trimmed)) { list.unshift(trimmed); saveOperators(list); }
  res.json({ ok: true, operators: loadOperators() });
});
app.delete('/operators/:name', (req, res) => {
  const list = loadOperators().filter(n => n !== decodeURIComponent(req.params.name));
  saveOperators(list);
  res.json({ ok: true, operators: list });
});

// Mode
app.post('/mode', (req, res) => {
  const { mode } = req.body;
  if (!['remote','streamdeck'].includes(mode)) return res.status(400).json({ error: 'Invalid mode' });
  state.mode = mode;
  broadcast({ type: 'mode_changed', mode });
  res.json({ ok: true, mode });
});

// ── WebSocket ─────────────────────────────────────────────────────────────────
wss.on('connection', (ws, req) => {
  ws.send(JSON.stringify({
    type: 'init',
    live: state.live,
    preview: state.preview,
    overlayLive: state.overlayLive,
    mode: state.mode,
    namesCount: namesDb.length,
    history: state.history.slice(0, 10),
    vmixConnected: state.vmixConnected,
    config,
    sessionStart: state.sessionStart,
    sessionLogCount: state.sessionLog.length
  }));
  ws.on('message', raw => {
    try { const m = JSON.parse(raw); if (m.type === 'ping') ws.send(JSON.stringify({ type: 'pong' })); } catch(e) {}
  });
});

// ── Start ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
server.listen(PORT, '0.0.0.0', () => {
  const nets = require('os').networkInterfaces();
  let localIp = 'localhost';
  for (const ifaces of Object.values(nets))
    for (const i of ifaces)
      if (i.family === 'IPv4' && !i.internal) { localIp = i.address; break; }

  console.log('\n╔══════════════════════════════════════════════╗');
  console.log('║   VMix Title Controller v2 — Running         ║');
  console.log('╠══════════════════════════════════════════════╣');
  console.log(`║  Local:    http://localhost:${PORT}              ║`);
  console.log(`║  Network:  http://${localIp}:${PORT}        ║`);
  console.log('╚══════════════════════════════════════════════╝\n');
});
