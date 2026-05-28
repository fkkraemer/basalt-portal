require('dotenv').config();
const express = require('express');
const fs      = require('fs');
const path    = require('path');
const XLSX    = require('xlsx');

const app  = express();
const PORT = process.env.PORT || 3030;

// Healthcheck for Railway (must be before Basic Auth)
app.get('/health', (req, res) => res.send('ok'));

// Passwortschutz (APP_PASSWORD als Umgebungsvariable setzen)
app.use((req, res, next) => {
  const pwd = process.env.APP_PASSWORD;
  if (!pwd) return next();
  const auth = req.headers.authorization || '';
  if (auth.startsWith('Basic ')) {
    const decoded = Buffer.from(auth.slice(6), 'base64').toString('utf8');
    if (decoded.split(':').slice(1).join(':') === pwd) return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="Basalt Portal"');
  res.status(401).send('Bitte Passwort eingeben.');
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Konstanten ───────────────────────────────────────────────────────────────

const TOKEN_URL = 'https://id.basalt.de/realms/basalt/protocol/openid-connect/token';
const CLIENT_ID = 'basalt-customer-portal';
const API_BASE  = 'https://apps.basalt.de/api';

// ─── State ────────────────────────────────────────────────────────────────────

let state = {
  phase: 'idle', logs: [], data: {}, error: null, lastRun: null, summary: null,
};

function log(msg) {
  const entry = { t: new Date().toLocaleTimeString('de-DE'), msg };
  state.logs.push(entry);
  console.log(`[${entry.t}] ${msg}`);
  if (state.logs.length > 200) state.logs.shift();
}

// ─── Sync-State ───────────────────────────────────────────────────────────────

function loadSyncState() {
  try { return JSON.parse(fs.readFileSync('sync-state.json', 'utf8')); }
  catch { return {}; }
}

function saveSyncState(s) {
  fs.writeFileSync('sync-state.json', JSON.stringify(s, null, 2));
}

function loadExistingData() {
  try { return JSON.parse(fs.readFileSync('output.json', 'utf8')); }
  catch { return {}; }
}

function mergeData(existing, fresh) {
  const merged = { ...existing };

  // Bestellungen: merge by order id
  const existingOrders   = existing.bestellungen || [];
  const existingOrderIds = new Set(existingOrders.map(o => o.id));
  const newOrders        = (fresh.bestellungen || []).filter(o => !existingOrderIds.has(o.id));
  merged.bestellungen    = [...existingOrders, ...newOrders];

  // Lieferscheine: merge by deliveryNoteNumber + materialNumber (Positionsschlüssel)
  const existingLS     = existing.lieferscheine || [];
  const existingLSKeys = new Set(existingLS.map(ls => `${ls.deliveryNoteNumber}|${ls.materialNumber || ls.material}`));
  const newLS          = (fresh.lieferscheine || []).filter(ls =>
    !existingLSKeys.has(`${ls.deliveryNoteNumber}|${ls.materialNumber || ls.material}`)
  );
  merged.lieferscheine = [...existingLS, ...newLS];

  // Kleine Stammdaten immer überschreiben
  if (fresh.lieferpunkte) merged.lieferpunkte = fresh.lieferpunkte;
  if (fresh.vertraege)    merged.vertraege    = fresh.vertraege;

  return merged;
}

// ─── Auth ─────────────────────────────────────────────────────────────────────

async function getToken(username, password) {
  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'password', client_id: CLIENT_ID, username, password, scope: 'openid' }).toString(),
  });
  if (!res.ok) throw new Error(`Token-Fehler (${res.status}): ${await res.text()}`);
  return (await res.json()).access_token;
}

// ─── API-Helfer ───────────────────────────────────────────────────────────────

async function apiGet(token, apiPath, params = {}) {
  const url = new URL(`${API_BASE}${apiPath}`);
  for (const [k, v] of Object.entries(params)) {
    if (v !== undefined && v !== null) url.searchParams.set(k, String(v));
  }
  const res = await fetch(url.toString(), { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`HTTP ${res.status} für ${url}`);
  return res.json();
}

// ─── Scraper ──────────────────────────────────────────────────────────────────

async function runScraper() {
  const username = process.env.BASALT_USERNAME;
  const password = process.env.BASALT_PASSWORD;
  if (!username || !password) throw new Error('Zugangsdaten fehlen');

  const sync    = loadSyncState();
  const isFirst = !sync.lastSyncDate;
  const fromDate = sync.lastSyncDate
    ? new Date(sync.lastSyncDate)
    : (() => { const d = new Date(); d.setFullYear(d.getFullYear() - 1); return d; })();

  log(isFirst
    ? 'Erster Abruf — lade letzte 12 Monate...'
    : `Inkrementeller Sync seit ${formatDate(fromDate.toISOString())}...`
  );

  log('Hole Token...');
  const token = await getToken(username, password);
  log('Token OK. Starte Abruf...');

  const results = {};

  // 1. Verträge
  log('Lade Verträge...');
  try {
    results.vertraege = await apiGet(token, '/contracts/view/customer-portal/contracts-overview');
    log(`Verträge: ${countRows(results.vertraege)} Einträge`);
  } catch (e) { log(`Verträge: ${e.message}`); }

  // 2. Lieferpunkte
  log('Lade Lieferpunkte...');
  let shippingPoints = [];
  try {
    const sp = await apiGet(token, '/scale-notes/view/customer-portal/contract-overview/shipping-points');
    results.lieferpunkte = sp;
    const spArray = sp.shippingPointsWithFactoryType || extractRows(sp);
    shippingPoints = spArray.map(r => r.shippingPointId || r.id).filter(Boolean);
    log(`Lieferpunkte: ${shippingPoints.join(', ')}`);
  } catch (e) { log(`Lieferpunkte: ${e.message}`); }

  // 3. Bestellungen (seit letztem Sync)
  log('Lade Bestellungen...');
  const alleBestellungen = [];
  const today = new Date();

  for (const spId of shippingPoints) {
    try {
      const filter = JSON.stringify({
        shippingPointId: spId,
        orderStartDate: { from: fromDate.toISOString(), to: today.toISOString() },
      });
      let page = 1;
      let totalOrders = null;

      while (true) {
        const data = await apiGet(token, '/orders/view/customer-portal/order-overview', {
          page, pageSize: 100,
          sortingField: 'deliveryStart', sortingDirection: 'desc',
          filter,
        });
        const rows = Array.isArray(data.items) ? data.items : extractRows(data);
        if (totalOrders === null) {
          totalOrders = data.filteredOrderCount ?? data.totalOrderCount ?? rows.length;
          log(`  ${spId}: ${totalOrders} Bestellungen`);
        }
        rows.forEach(r => { r._shippingPoint = spId; });
        alleBestellungen.push(...rows);
        const fetched = alleBestellungen.filter(r => r._shippingPoint === spId).length;
        if (rows.length < 100 || fetched >= totalOrders) break;
        page++;
      }
    } catch (e) { log(`  ${spId} Bestellungen: ${e.message}`); }
  }
  results.bestellungen = alleBestellungen;
  log(`Bestellungen: ${alleBestellungen.length} geladen`);

  // 4. Lieferscheine (neueste zuerst, stopp wenn älter als letzter Sync)
  log('Lade Lieferscheine...');
  const alleLieferscheine = [];

  for (const spId of shippingPoints) {
    try {
      let page = 1;

      while (true) {
        const data = await apiGet(token, `/scale-notes/view/customer-portal/factory/${spId}/delivery-overview`, {
          page, pageSize: 100,
          sortingField: 'deliveryNoteNumber', sortingDirection: 'desc',
        });
        const groups = Array.isArray(data.items) ? data.items : [];
        const raw    = groups.flat(2);
        raw.forEach(r => { r._shippingPoint = spId; });

        if (!isFirst) {
          // Nur neue Einträge — stopp wenn Datum vor letztem Sync
          const neuePositionen = raw.filter(r =>
            r.deliveryNoteCreationDate && new Date(r.deliveryNoteCreationDate) >= fromDate
          );
          alleLieferscheine.push(...neuePositionen);
          // Wenn älteste Position auf dieser Seite schon vor fromDate → fertig
          const oldest = raw.reduce((min, r) =>
            r.deliveryNoteCreationDate < min ? r.deliveryNoteCreationDate : min,
            raw[0]?.deliveryNoteCreationDate || ''
          );
          if (!oldest || new Date(oldest) < fromDate) break;
        } else {
          alleLieferscheine.push(...raw);
        }

        if (groups.length < 100) break;
        page++;
      }
      log(`  ${spId}: ${alleLieferscheine.filter(r => r._shippingPoint === spId).length} Lieferscheine`);
    } catch (e) { log(`  ${spId} Lieferscheine: ${e.message}`); }
  }
  results.lieferscheine = alleLieferscheine;
  log(`Lieferscheine: ${alleLieferscheine.length} geladen`);

  return results;
}

// ─── Excel-Export ─────────────────────────────────────────────────────────────

function generateExcel(data) {
  const wb           = XLSX.utils.book_new();
  const bestellungen = data.bestellungen  || [];
  const lieferscheine = data.lieferscheine || [];

  // Blatt 1: Rechnungsabgleich (pro Lieferschein)
  addSheet(wb, 'Rechnungsabgleich', buildRechnungsabgleich(lieferscheine));

  // Blatt 2: Bestellung vs. Lieferung
  addSheet(wb, 'Bestellung vs. Lieferung', buildAbgleich(bestellungen, lieferscheine));

  // Blatt 3: Lieferscheine (Rohdaten, nicht storniert)
  addSheet(wb, 'Lieferscheine', lieferscheine
    .filter(ls => !ls.cancelled)
    .sort((a, b) => (b.deliveryNoteNumber || '').localeCompare(a.deliveryNoteNumber || ''))
    .map(ls => ({
      'Lieferschein-Nr':  ls.deliveryNoteNumber,
      'Datum':            formatDate(ls.deliveryNoteCreationDate),
      'Lieferpunkt':      ls._shippingPoint || ls.shippingPointId,
      'Baustelle':        ls.constructionSite || ls.destination1 || '',
      'Material':         ls.materialName || ls.material,
      'Menge':            ls.amount,
      'Einheit':          ls.unit,
      'Vertragsnummer':   ls.contractNumber,
      'Incoterm':         ls.incoterm || '',
      'Storniert':        ls.cancelled ? 'Ja' : 'Nein',
    }))
  );

  // Blatt 4: Bestellungen
  addSheet(wb, 'Bestellungen', bestellungen
    .sort((a, b) => (b.deliveryStart || '').localeCompare(a.deliveryStart || ''))
    .map(b => ({
      'Bestell-ID':       b.id,
      'Status':           translateStatus(b.state),
      'Lieferdatum':      formatDate(b.deliveryStart),
      'Lieferpunkt':      b.shippingPoint || b._shippingPoint,
      'Material':         b.materialName,
      'Bestellt (Menge)': b.amount,
      'Einheit':          b.unit,
      'Baustelle':        b.constructionSite || '',
      'Vertragsnummer':   b.contractNumber,
    }))
  );

  const filePath = path.resolve('basalt-report.xlsx');
  XLSX.writeFile(wb, filePath);
  log(`Excel gespeichert: basalt-report.xlsx`);
  return filePath;
}

function buildRechnungsabgleich(lieferscheine) {
  const groups = {};
  for (const ls of lieferscheine) {
    if (ls.cancelled) continue;
    const key = ls.deliveryNoteNumber;
    if (!groups[key]) {
      groups[key] = {
        'Lieferschein-Nr': ls.deliveryNoteNumber,
        'Datum':           formatDate(ls.deliveryNoteCreationDate),
        'Lieferpunkt':     ls._shippingPoint || ls.shippingPointId,
        'Baustelle':       ls.constructionSite || ls.destination1 || '',
        'Vertragsnummer':  ls.contractNumber,
        positionen: [],
      };
    }
    groups[key].positionen.push(`${ls.materialName || ls.material}: ${ls.amount} ${ls.unit}`);
  }
  return Object.values(groups)
    .sort((a, b) => b['Lieferschein-Nr'].localeCompare(a['Lieferschein-Nr']))
    .map(g => ({
      'Lieferschein-Nr': g['Lieferschein-Nr'],
      'Datum':           g.Datum,
      'Lieferpunkt':     g.Lieferpunkt,
      'Baustelle':       g.Baustelle,
      'Vertragsnummer':  g.Vertragsnummer,
      'Materialien / Mengen': g.positionen.join(' | '),
    }));
}

function buildAbgleich(bestellungen, lieferscheine) {
  const groups = {};

  for (const b of bestellungen) {
    const key = `${b.contractNumber}||${b.materialName}`;
    if (!groups[key]) groups[key] = {
      'Vertragsnummer':   b.contractNumber,
      'Baustelle':        b.constructionSite || '',
      'Material':         b.materialName,
      'Einheit':          b.unit,
      _bestellt:          0,
      _geliefert:         0,
    };
    groups[key]._bestellt += b.amount || 0;
  }

  for (const ls of lieferscheine) {
    if (ls.cancelled) continue;
    const key = `${ls.contractNumber}||${ls.materialName || ls.material}`;
    if (!groups[key]) groups[key] = {
      'Vertragsnummer':   ls.contractNumber,
      'Baustelle':        ls.constructionSite || ls.destination1 || '',
      'Material':         ls.materialName || ls.material,
      'Einheit':          ls.unit,
      _bestellt:          0,
      _geliefert:         0,
    };
    groups[key]._geliefert += ls.amount || 0;
  }

  return Object.values(groups).map(g => {
    const bestellt   = round2(g._bestellt);
    const geliefert  = round2(g._geliefert);
    const differenz  = round2(bestellt - geliefert);
    return {
      'Vertragsnummer':     g.Vertragsnummer,
      'Baustelle':          g.Baustelle,
      'Material':           g.Material,
      'Bestellt (Menge)':   bestellt,
      'Geliefert (Menge)':  geliefert,
      'Einheit':            g.Einheit,
      'Differenz':          differenz,
      'Status':             bestellt === 0  ? 'Nur Lieferung' :
                            geliefert === 0 ? 'Ausstehend' :
                            geliefert >= bestellt ? 'Vollständig' : 'Teillieferung',
    };
  }).sort((a, b) => (a.Vertragsnummer || '').localeCompare(b.Vertragsnummer || ''));
}

function addSheet(wb, name, rows) {
  if (!rows.length) return;
  const ws = XLSX.utils.json_to_sheet(rows);
  const headers = Object.keys(rows[0]);
  ws['!cols'] = headers.map(h => ({ wch: Math.min(Math.max(h.length + 2, 14), 50) }));
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  XLSX.utils.book_append_sheet(wb, ws, name);
}

// ─── Hilfsf. ──────────────────────────────────────────────────────────────────

function formatDate(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }); }
  catch { return iso; }
}

function translateStatus(s) {
  return { COMPLETED: 'Abgeschlossen', OPEN: 'Offen', IN_PROGRESS: 'In Bearbeitung', CANCELLED: 'Storniert' }[s] || s || '';
}

function extractRows(data) {
  if (Array.isArray(data)) return data;
  for (const k of ['items', 'data', 'results', 'records', 'rows', 'list', 'content']) {
    if (Array.isArray(data[k])) return data[k];
  }
  return [data];
}

function countRows(data) { return data ? extractRows(data).length : 0; }
function round2(n) { return Math.round((n || 0) * 100) / 100; }

// ─── API-Routen ───────────────────────────────────────────────────────────────

app.get('/api/credentials', (req, res) => {
  res.json({ username: process.env.BASALT_USERNAME || '', saved: !!(process.env.BASALT_USERNAME && process.env.BASALT_PASSWORD) });
});

app.post('/api/credentials', (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: 'Beide Felder erforderlich' });
  process.env.BASALT_USERNAME = username;
  process.env.BASALT_PASSWORD = password;
  fs.writeFileSync('.env', `BASALT_USERNAME=${username}\nBASALT_PASSWORD=${password}\n`, 'utf8');
  res.json({ ok: true });
});

app.get('/api/status', (req, res) => {
  const sync = loadSyncState();
  res.json({
    phase:     state.phase,
    logs:      state.logs.slice(-30),
    hasData:   Object.keys(state.data).length > 0,
    lastRun:   state.lastRun,
    lastSync:  sync.lastSyncDate ? formatDate(sync.lastSyncDate) : null,
    summary:   state.summary,
    error:     state.error,
  });
});

app.post('/api/run', (req, res) => {
  if (state.phase === 'running') return res.status(400).json({ error: 'Läuft bereits' });
  state = { phase: 'running', logs: [], data: {}, error: null, lastRun: null, summary: null };
  res.json({ ok: true });

  runScraper()
    .then(fresh => {
      const existing = loadExistingData();
      const merged   = mergeData(existing, fresh);
      state.data     = merged;

      fs.writeFileSync('output.json', JSON.stringify(merged, null, 2));

      const newOrders = (fresh.bestellungen || []).length;
      const newLS     = (fresh.lieferscheine || []).length;
      log(`Neu: ${newOrders} Bestellungen, ${newLS} Lieferscheine`);
      log(`Gesamt: ${merged.bestellungen?.length || 0} Bestellungen, ${merged.lieferscheine?.length || 0} Lieferscheine`);

      generateExcel(merged);
      saveSyncState({ lastSyncDate: new Date().toISOString() });

      state.phase   = 'done';
      state.lastRun = new Date().toLocaleString('de-DE');
      state.summary = {
        bestellungenGesamt:   merged.bestellungen?.length  || 0,
        lieferscheineGesamt:  merged.lieferscheine?.length || 0,
        neueBestellungen:     newOrders,
        neueLieferscheine:    newLS,
      };
      log('Fertig!');
    })
    .catch(e => {
      state.phase = 'error';
      state.error = e.message;
      log('Fehler: ' + e.message);
    });
});

app.get('/api/data', (req, res) => res.json(state.data));

app.get('/api/config', (req, res) => res.json({ orderEmail: process.env.ORDER_EMAIL || '' }));

app.get('/api/download/excel', (req, res) => {
  const file = path.resolve('basalt-report.xlsx');
  if (!fs.existsSync(file)) return res.status(404).json({ error: 'Noch kein Report vorhanden' });
  res.download(file, `Basalt-Report-${new Date().toISOString().slice(0, 10)}.xlsx`);
});

// ─── Start ────────────────────────────────────────────────────────────────────

app.listen(PORT, '0.0.0.0', () => {
  console.log(`\nBasalt Portal Scraper → http://localhost:${PORT}\n`);
});
