// ============================================================
// Dvůr Pod Dubem – api2.js (ROW-BASED backend, fáze 1)
// Nový datový model: jeden list "Objednávky", řádek = objednávka.
// Sloupce: A:ID | B:Datum | C:Jméno | D:Produkt | E:Kusy | F:Stav | G:RequestID | H:Vytvořeno
// Stav = text 'čeká' / 'doručeno' (barvy už nejsou nositelem stavu).
// RequestID váže řádek na Požadavky → přesná synchronizace úprav.
//
// TESTOVACÍ REŽIM: míří na kopii sheetu, push vypnutý.
// Před ostrým nasazením: přepnout SPREADSHEET_ID na LIVE a PUSH_ENABLED na true.
// ============================================================

const { google } = require('googleapis');

// ── PŘEPÍNAČE PROSTŘEDÍ ──
const SPREADSHEET_ID_LIVE = '1ORfd4FhxKsJIuk22WvdoW6sLNpRNBFK39q3yt5YAj8Q';
const SPREADSHEET_ID_TEST = '1CEG_-bfKRBTvNpzeZDZsN9D_HHbaH9m8USwrTKOR9gw';
const SPREADSHEET_ID = SPREADSHEET_ID_TEST;   // ← fáze 1–3: TEST. Ostrý přechod: LIVE.
const PUSH_ENABLED = false;                    // ← fáze 1–3: false. Ostrý přechod: true.

const ORDERS_SHEET = 'Objednávky';
const ORDERS_HDR = ['ID','Datum','Jméno','Produkt','Kusy','Stav','RequestID','Vytvořeno'];
const DISCOUNT_NAMES = ['Danča', 'Kristýna', 'Renata', 'Andrea', 'Marcela', 'Dáša'];
const HEADER_ROW = 2;      // pro čtení STARÝCH listů při migraci
const FIRST_COL = 2;
const DATA_START_ROW = 4;

const VAPID_PUBLIC = process.env.VAPID_PUBLIC || '';
const VAPID_PRIVATE = process.env.VAPID_PRIVATE || '';
const VAPID_SUBJECT = 'mailto:dvurpoddubem@email.cz';

async function getSheets() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  return google.sheets({ version: 'v4', auth });
}

function colLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function fmtDate(val) {
  if (!val) return '';
  const s = val.toString().trim();
  const m0 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m0) return parseInt(m0[1]) + '.' + parseInt(m0[2]) + '.' + m0[3]; // kanonicky bez úvodních nul
  if (s.match(/^\d+$/) && parseInt(s) > 1000) {
    const date = new Date((parseInt(s) - 25569) * 86400 * 1000);
    if (!isNaN(date)) return date.getUTCDate() + '.' + (date.getUTCMonth() + 1) + '.' + date.getUTCFullYear();
  }
  const d = new Date(s);
  if (!isNaN(d)) return d.getUTCDate() + '.' + (d.getUTCMonth() + 1) + '.' + d.getUTCFullYear();
  return s;
}

function parseDate(val) {
  if (!val) return null;
  const s = val.toString().trim();
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  if (s.match(/^\d+$/) && parseInt(s) > 1000) return new Date((parseInt(s) - 25569) * 86400 * 1000);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function dnesCZ() {
  const dn = new Date();
  return dn.getDate() + '.' + (dn.getMonth() + 1) + '.' + dn.getFullYear();
}

async function getSheetData(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:Z`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  return res.data.values || [];
}

async function getSheetId(sheets, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  return res.data.sheets.find(s => s.properties.title === sheetName)?.properties?.sheetId;
}

// ════════════════════════════════════════════════════
// OBJEDNÁVKY (nový řádkový model)
// ════════════════════════════════════════════════════

async function getOrdersData(sheets) {
  try { return await getSheetData(sheets, ORDERS_SHEET); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: ORDERS_SHEET } } }] } });
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `${ORDERS_SHEET}!A1:H1`, valueInputOption: 'USER_ENTERED', requestBody: { values: [ORDERS_HDR] } });
    return [ORDERS_HDR];
  }
}

// Mapování řádku na objekt (index = pozice v datech, radek = 1-indexed v sheetu)
function mapOrders(data) {
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i] || [];
    if (r[0] === '' || r[0] === null || r[0] === undefined) continue;
    out.push({
      radek: i + 1,
      id: parseInt(r[0]) || 0,
      datum: fmtDate(r[1]),
      jmeno: (r[2] || '').toString(),
      produkt: (r[3] || '').toString().trim().toLowerCase(),
      kusy: parseFloat((r[4] || '0').toString().replace(',', '.')) || 0,
      stav: (r[5] || '').toString().trim().toLowerCase(),
      requestId: (r[6] || '').toString(),
    });
  }
  return out;
}

function nextOrderId(orders) {
  let max = 0;
  orders.forEach(o => { if (o.id > max) max = o.id; });
  return max + 1;
}

// Přidá jednu objednávku jako řádek. Vrací její ID.
async function pridejObjednavku(sheets, jmeno, kusy, datum, produkt, requestId) {
  const data = await getOrdersData(sheets);
  const orders = mapOrders(data);
  const id = nextOrderId(orders);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID, range: `${ORDERS_SHEET}!A:H`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[id, datum, jmeno, produkt, kusy, 'čeká', requestId || '', dnesCZ()]] }
  });
  return id;
}

// Najde řádek objednávky podle ID. Vrací {radek, order} nebo null.
async function najdiObjednavku(sheets, id) {
  const data = await getOrdersData(sheets);
  const orders = mapOrders(data);
  const o = orders.find(x => x.id === parseInt(id));
  return o ? o : null;
}

// Ceník
function cenaZa(jmeno, produkt, kusy) {
  if (produkt === 'vajicka') return kusy * (DISCOUNT_NAMES.includes(jmeno) ? 9 : 10);
  if (produkt === 'bedynka') return kusy * 490;
  return Math.round(kusy * 190); // sirup
}

// ════════════════════════════════════════════════════
// POŽADAVKY (beze změny modelu, sloupce A–L)
// ════════════════════════════════════════════════════

function mapRequests(data) {
  return data.length <= 1 ? [] : data.slice(1).map((r, i) => ({
    radek: i + 2, id: r[0], jmeno: r[1], datumPozadavku: fmtDate(r[2]), datumDoruceni: fmtDate(r[3]),
    vajicka: r[4] || 0, bedynka: r[5] || 0, sirup: r[6] || 0, celkem: r[7] || 0,
    stav: r[8] || 'čeká na potvrzení', navrzenyTermin: fmtDate(r[9]), zprava: (r[10] || '').toString(),
    zaplaceno: (r[11] || '').toString().trim().toLowerCase() === 'ano'
  }));
}

async function getPozadavkyData(sheets) {
  try { return await getSheetData(sheets, 'Požadavky'); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Požadavky' } } }] } });
    const hdr = [['ID','Zákazník','Datum požadavku','Datum doručení','Vajíčka','Bedýnka','Sirup','Celkem','Stav','Navržený termín','Zpráva','Zaplaceno']];
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: 'Požadavky!A1:L1', valueInputOption: 'USER_ENTERED', requestBody: { values: hdr } });
    return hdr;
  }
}

async function getZakazniciData(sheets) {
  try { return await getSheetData(sheets, 'Zákazníci'); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Zákazníci' } } }] } });
    const rows = [['Jméno','PIN','Datum registrace'], ...['Danča','Kristýna','Renata','Andrea','Marcela','Dáša','Regina','Ondra'].map(j => [j,'',''])];
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: 'Zákazníci!A1:C' + rows.length, valueInputOption: 'USER_ENTERED', requestBody: { values: rows } });
    return rows;
  }
}

// ════════════════════════════════════════════════════
// SKLAD / AKTUALITY (převzato beze změny)
// ════════════════════════════════════════════════════

const SKLAD_DEFAULT = [
  ['Produkt','Dostupné','K dispozici od','Popis'],
  ['sirup','ano','','malinový'],
  ['bedynka','ne','','Mrkev, brambory, cuketa, salát, bylinky'],
  ['vajicka','ano','','']
];

async function getSkladData(sheets) {
  let data;
  try { data = await getSheetData(sheets, 'Sklad'); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Sklad' } } }] } });
    data = [];
  }
  if (!data || data.length === 0 || !data[0] || (data[0][0]||'').toString().toLowerCase().indexOf('produkt') < 0) {
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: 'Sklad!A1:D' + SKLAD_DEFAULT.length, valueInputOption: 'USER_ENTERED', requestBody: { values: SKLAD_DEFAULT } });
    data = SKLAD_DEFAULT.map(r => r.slice());
  }
  return data;
}

function mapSklad(data) {
  const out = { sirup:{}, bedynka:{}, vajicka:{} };
  for (let i = 1; i < data.length; i++) {
    const key = (data[i][0]||'').toString().trim().toLowerCase();
    if (!out[key]) continue;
    out[key] = {
      dostupne: (data[i][1]||'').toString().trim().toLowerCase() === 'ano',
      od: fmtDate(data[i][2]),
      popis: (data[i][3]||'').toString().trim()
    };
  }
  return out;
}

async function setSklad(sheets, produkt, dostupne, od, popis) {
  const data = await getSkladData(sheets);
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0]||'').toString().trim().toLowerCase() === produkt) { row = i + 1; break; }
  }
  if (row === -1) {
    await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: 'Sklad!A:D', valueInputOption: 'USER_ENTERED', requestBody: { values: [[produkt, dostupne?'ano':'ne', od||'', popis||'']] } });
  } else {
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Sklad!B${row}:D${row}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[dostupne?'ano':'ne', od||'', popis||'']] } });
  }
}

async function getAktualitaData(sheets) {
  try { return await getSheetData(sheets, 'Aktuality'); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Aktuality' } } }] } });
    const hdr = [['Datum','Nadpis','Text','URL_fotky','Aktivní']];
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: 'Aktuality!A1:E1', valueInputOption: 'USER_ENTERED', requestBody: { values: hdr } });
    return hdr;
  }
}

async function getAktualita(sheets) {
  const data = await getAktualitaData(sheets);
  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i] || [];
    const aktivni = (r[4] || '').toString().trim().toLowerCase();
    if (aktivni === 'ne') continue;
    const nadpis = (r[1] || '').toString().trim();
    const text = (r[2] || '').toString().trim();
    const foto = (r[3] || '').toString().trim();
    if (!nadpis && !text && !foto) continue;
    return { datum: fmtDate(r[0]), nadpis, text, foto };
  }
  return null;
}

async function setAktualita(sheets, nadpis, text, fotoUrl) {
  await getAktualitaData(sheets);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID, range: 'Aktuality!A:E',
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[dnesCZ(), nadpis || '', text || '', fotoUrl || '', 'ano']] }
  });
}

async function clearAktualita(sheets) {
  const data = await getAktualitaData(sheets);
  const updates = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i] || [];
    if ((r[4] || '').toString().trim().toLowerCase() !== 'ne') {
      updates.push({ range: 'Aktuality!E' + (i + 1), values: [['ne']] });
    }
  }
  if (updates.length) {
    await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { valueInputOption: 'USER_ENTERED', data: updates } });
  }
  return updates.length;
}

async function editAktualita(sheets, nadpis, text, fotoUrl) {
  const data = await getAktualitaData(sheets);
  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i] || [];
    if ((r[4] || '').toString().trim().toLowerCase() === 'ne') continue;
    const hasContent = (r[1] || '').toString().trim() || (r[2] || '').toString().trim() || (r[3] || '').toString().trim();
    if (!hasContent) continue;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID, range: 'Aktuality!B' + (i + 1) + ':D' + (i + 1),
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[nadpis || '', text || '', fotoUrl || '']] }
    });
    return true;
  }
  return false;
}

// ════════════════════════════════════════════════════
// GRAIN (náklady) – beze změny
// ════════════════════════════════════════════════════

function normHdr(s) {
  return (s || '').toString().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function findGrainData(data) {
  let dCol = -1, cCol = -1, kCol = -1, hdrRow = -1;
  const maxScan = Math.min(data.length, 12);
  for (let r = 0; r < maxScan && hdrRow === -1; r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      const h = normHdr(row[c]);
      if (h === 'datum objednavky') dCol = c;
      if (h === 'cena zrni') cCol = c;
      if (h === 'kg zrni') kCol = c;
    }
    if (dCol >= 0 && cCol >= 0) hdrRow = r;
    else { dCol = -1; cCol = -1; kCol = -1; }
  }
  if (hdrRow === -1) return [];
  const out = [];
  for (let r = hdrRow + 1; r < data.length; r++) {
    const row = data[r] || [];
    if (!row[dCol] || row[cCol] === '' || row[cCol] === null || row[cCol] === undefined) continue;
    out.push({ datum: row[dCol], cena: row[cCol], kg: kCol >= 0 ? (row[kCol] ?? '') : '' });
  }
  return out;
}

// ════════════════════════════════════════════════════
// WEB PUSH (v testu vypnutý přes PUSH_ENABLED)
// ════════════════════════════════════════════════════
const crypto = require('crypto');

function b64urlEncode(buf) {
  return Buffer.from(buf).toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
}
function b64urlDecode(str) {
  str = str.replace(/-/g,'+').replace(/_/g,'/');
  while (str.length % 4) str += '=';
  return Buffer.from(str, 'base64');
}

function vapidJWT(audience) {
  const header = b64urlEncode(JSON.stringify({ typ: 'JWT', alg: 'ES256' }));
  const payload = b64urlEncode(JSON.stringify({ aud: audience, exp: Math.floor(Date.now()/1000) + 12*3600, sub: VAPID_SUBJECT }));
  const unsigned = header + '.' + payload;
  const privBuf = b64urlDecode(VAPID_PRIVATE);
  const pubBuf = b64urlDecode(VAPID_PUBLIC);
  const jwk = { kty: 'EC', crv: 'P-256', d: b64urlEncode(privBuf), x: b64urlEncode(pubBuf.slice(1, 33)), y: b64urlEncode(pubBuf.slice(33, 65)) };
  const keyObj = crypto.createPrivateKey({ key: jwk, format: 'jwk' });
  const der = crypto.sign('sha256', Buffer.from(unsigned), { key: keyObj, dsaEncoding: 'ieee-p1363' });
  return unsigned + '.' + b64urlEncode(der);
}

function hkdf(salt, ikm, info, length) {
  const prk = crypto.createHmac('sha256', salt).update(ikm).digest();
  const out = crypto.createHmac('sha256', prk).update(Buffer.concat([info, Buffer.from([1])])).digest();
  return out.slice(0, length);
}

function encryptPayload(payloadStr, p256dhB64, authB64) {
  const clientPub = b64urlDecode(p256dhB64);
  const auth = b64urlDecode(authB64);
  const salt = crypto.randomBytes(16);
  const ecdh = crypto.createECDH('prime256v1');
  ecdh.generateKeys();
  const serverPub = ecdh.getPublicKey();
  const sharedSecret = ecdh.computeSecret(clientPub);
  const authInfo = Buffer.concat([Buffer.from('WebPush: info\0'), clientPub, serverPub]);
  const ikm = hkdf(auth, sharedSecret, authInfo, 32);
  const cek = hkdf(salt, ikm, Buffer.from('Content-Encoding: aes128gcm\0'), 16);
  const nonce = hkdf(salt, ikm, Buffer.from('Content-Encoding: nonce\0'), 12);
  const padded = Buffer.concat([Buffer.from(payloadStr, 'utf8'), Buffer.from([0x02])]);
  const cipher = crypto.createCipheriv('aes-128-gcm', cek, nonce);
  const encrypted = Buffer.concat([cipher.update(padded), cipher.final(), cipher.getAuthTag()]);
  const header = Buffer.concat([salt, Buffer.from([0,0,0x10,0x00]), Buffer.from([serverPub.length]), serverPub]);
  return Buffer.concat([header, encrypted]);
}

function sendOnePush(sub, payloadStr) {
  return new Promise((resolve) => {
    try {
      if (!VAPID_PUBLIC || !VAPID_PRIVATE) return resolve({ ok:false, gone:false });
      const url = new URL(sub.endpoint);
      const body = encryptPayload(payloadStr, sub.p256dh, sub.auth);
      const jwt = vapidJWT(url.origin);
      const https = require('https');
      const req = https.request({
        hostname: url.hostname, path: url.pathname + url.search, method: 'POST',
        headers: {
          'Content-Type': 'application/octet-stream', 'Content-Encoding': 'aes128gcm',
          'Content-Length': body.length, 'TTL': '86400', 'Urgency': 'high',
          'Authorization': 'vapid t=' + jwt + ', k=' + VAPID_PUBLIC
        }
      }, (res) => {
        const gone = res.statusCode === 404 || res.statusCode === 410;
        res.on('data', ()=>{}); res.on('end', ()=> resolve({ ok: res.statusCode>=200&&res.statusCode<300, gone }));
      });
      req.on('error', () => resolve({ ok:false, gone:false }));
      req.write(body); req.end();
    } catch(e) { resolve({ ok:false, gone:false }); }
  });
}

async function getPushData(sheets) {
  try { return await getSheetData(sheets, 'Push'); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Push' } } }] } });
    const hdr = [['Jméno','Endpoint','p256dh','auth']];
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: 'Push!A1:D1', valueInputOption: 'USER_ENTERED', requestBody: { values: hdr } });
    return hdr;
  }
}

async function savePushSub(sheets, jmeno, endpoint, p256dh, auth) {
  const data = await getPushData(sheets);
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === endpoint) {
      await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Push!A${i+1}:D${i+1}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[jmeno, endpoint, p256dh, auth]] } });
      return;
    }
  }
  await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: 'Push!A:D', valueInputOption: 'USER_ENTERED', requestBody: { values: [[jmeno, endpoint, p256dh, auth]] } });
}

async function getSubsFor(sheets, jmeno) {
  const data = await getPushData(sheets);
  const subs = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (jmeno && data[i][0] !== jmeno) continue;
    subs.push({ row: i+1, jmeno: data[i][0], endpoint: data[i][1], p256dh: data[i][2], auth: data[i][3] });
  }
  return subs;
}

async function pushToCustomer(sheets, jmeno, title, body, tag) {
  if (!PUSH_ENABLED) { console.log('[TEST] push potlačen: ' + (jmeno||'všem') + ' – ' + title + ' / ' + body); return; }
  if (!VAPID_PUBLIC) return;
  const subs = await getSubsFor(sheets, jmeno);
  const targetUrl = (jmeno === '__admin__') ? '/' : '/zakaznik';
  const payload = JSON.stringify({ title, body, tag: tag||'dpd', url: targetUrl });
  const goneRows = [];
  for (const s of subs) {
    const r = await sendOnePush(s, payload);
    if (r.gone) goneRows.push(s.row);
  }
  for (const row of goneRows.sort((a,b)=>b-a)) {
    try { await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: `Push!A${row}:D${row}` }); } catch(e){}
  }
}

// ════════════════════════════════════════════════════
// MIGRACE ze starého sloupcového modelu
// Čte Vajíčka/Bedýnky/Sirup, stav bere z barev buněk
// (oranžová=čeká, zelená=doručeno). Dry-run defaultně.
// ════════════════════════════════════════════════════

async function getSheetColors(sheets, sheetName) {
  const res = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID, ranges: [`${sheetName}!A:Z`],
    fields: 'sheets.data.rowData.values.userEnteredFormat.backgroundColor'
  });
  return res.data.sheets?.[0]?.data?.[0]?.rowData || [];
}

function isOrange(bg) {
  return !!(bg && bg.red > 0.9 && bg.green > 0.6 && bg.green < 0.72 && bg.blue < 0.2);
}
function isGreen(bg) {
  return !!(bg && bg.red > 0.45 && bg.red < 0.7 && bg.green > 0.7 && bg.green < 0.85 && bg.blue > 0.38 && bg.blue < 0.6);
}

function getOldCustomers(headers) {
  const c = [];
  for (let i = FIRST_COL - 1; i < headers.length; i++) {
    if (headers[i] && headers[i] !== '') c.push({ jmeno: headers[i], col: i });
    else break;
  }
  return c;
}

async function migrace(sheets, zapis) {
  const produkty = [
    { key: 'vajicka', sheet: 'Vajíčka' },
    { key: 'bedynka', sheet: 'Bedýnky' },
    { key: 'sirup',   sheet: 'Sirup' }
  ];
  const radky = []; // {datum, jmeno, produkt, kusy, stav}
  const varovani = [];

  for (const p of produkty) {
    let data, colors;
    try {
      [data, colors] = await Promise.all([getSheetData(sheets, p.sheet), getSheetColors(sheets, p.sheet)]);
    } catch(e) { varovani.push(p.sheet + ': list nenalezen, přeskočeno'); continue; }
    const customers = getOldCustomers(data[HEADER_ROW - 1] || []);
    for (let i = DATA_START_ROW - 1; i < data.length; i++) {
      const dc = data[i]?.[0]; if (!dc) continue;
      if (typeof dc === 'string' && isNaN(Date.parse(dc)) && !dc.match(/^\d/)) continue; // "Květen 2026" apod.
      const ds = fmtDate(dc); if (!ds) continue;
      for (const c of customers) {
        const val = data[i][c.col];
        if (!val || val === '') continue;
        const kusy = parseFloat(val.toString().replace(/\s/g,'').replace(',','.'));
        if (!kusy || kusy <= 0) continue;
        const bg = colors[i]?.values?.[c.col]?.userEnteredFormat?.backgroundColor;
        let stav = null;
        if (isGreen(bg)) stav = 'doručeno';
        else if (isOrange(bg)) stav = 'čeká';
        else {
          // Bez stavové barvy: staré objednávky z doby před barevným systémem.
          // Datum v minulosti → považuj za doručené (jinak by zmizely z tržeb).
          const dObj = parseDate(ds);
          const dnes = new Date(); dnes.setHours(0,0,0,0);
          if (dObj && dObj.getTime() < dnes.getTime()) {
            stav = 'doručeno';
            varovani.push(p.sheet + ' ' + colLetter(c.col+1) + (i+1) + ': bez stavové barvy, datum v minulosti → převedeno jako doručeno (' + c.jmeno + ', ' + ds + ', ' + kusy + ')');
          } else {
            varovani.push(p.sheet + ' ' + colLetter(c.col+1) + (i+1) + ': bez stavové barvy a datum není v minulosti (' + c.jmeno + ', ' + ds + ', ' + kusy + ') – přeskočeno');
            continue;
          }
        }
        radky.push({ datum: ds, jmeno: c.jmeno, produkt: p.key, kusy, stav });
      }
    }
  }

  // Seřaď chronologicky podle data doručení
  radky.sort((a, b) => {
    const da = parseDate(a.datum), db = parseDate(b.datum);
    return (da?da.getTime():0) - (db?db.getTime():0);
  });

  // Best-effort přiřazení RequestID z Požadavků (potvrzeno + shoda jméno/datum/produkt/množství)
  let pozadavky = [];
  try { pozadavky = mapRequests(await getPozadavkyData(sheets)); } catch(e){}
  const pouzite = new Set();
  radky.forEach(r => {
    const kand = pozadavky.filter(q =>
      q.stav === 'potvrzeno' && !pouzite.has(q.id) &&
      q.jmeno === r.jmeno && q.datumDoruceni === r.datum &&
      Math.abs((parseFloat((q[{vajicka:'vajicka',bedynka:'bedynka',sirup:'sirup'}[r.produkt]]||'0').toString().replace(',','.'))||0) - r.kusy) < 0.001
    );
    if (kand.length === 1) { r.requestId = kand[0].id; pouzite.add(kand[0].id); }
    else r.requestId = '';
  });

  const souhrn = {
    celkem: radky.length,
    ceka: radky.filter(r=>r.stav==='čeká').length,
    doruceno: radky.filter(r=>r.stav==='doručeno').length,
    sRequestId: radky.filter(r=>r.requestId!=='').length,
    varovani
  };

  if (!zapis) {
    return { rezim: 'dry-run (nic nezapsáno)', souhrn, radky };
  }

  // Zápis: založ list (s hlavičkou), pak přidej všechny řádky s postupnými ID
  const existing = await getOrdersData(sheets);
  const existingOrders = mapOrders(existing);
  if (existingOrders.length > 0) {
    return { rezim: 'PŘERUŠENO', duvod: 'List ' + ORDERS_SHEET + ' už obsahuje ' + existingOrders.length + ' objednávek. Migrace se nespustí dvakrát – nejdřív list vyprázdni (kromě hlavičky).', souhrn };
  }
  let id = 1;
  const values = radky.map(r => [id++, r.datum, r.jmeno, r.produkt, r.kusy, r.stav, r.requestId, 'migrace ' + dnesCZ()]);
  if (values.length) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: `${ORDERS_SHEET}!A:H`,
      valueInputOption: 'USER_ENTERED', requestBody: { values }
    });
  }
  // Malý přehled vpravo (SUMIFS drží i po mazání řádků)
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: [
      { range: `${ORDERS_SHEET}!J1`, values: [['Přehled (ks/l celkem)']] },
      { range: `${ORDERS_SHEET}!J2:K4`, values: [
        ['Vajíčka', '=SUMIF(D:D;"vajicka";E:E)'],
        ['Bedýnky', '=SUMIF(D:D;"bedynka";E:E)'],
        ['Sirup',   '=SUMIF(D:D;"sirup";E:E)']
      ]}
    ]}
  });
  return { rezim: 'zápis proveden', zapsano: values.length, souhrn };
}

// ════════════════════════════════════════════════════
// HANDLER
// ════════════════════════════════════════════════════

const hdrs = { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' };

exports.handler = async function(event) {
  const p = event.queryStringParameters || {};
  try {
    const sheets = await getSheets();
    let result;
    switch (p.action) {

      case 'migrace': {
        const zapis = p.zapis === '1' || p.zapis === 'true';
        result = await migrace(sheets, zapis);
        break;
      }

      case 'add': {
        // Admin přidá objednávku ručně (bez požadavku)
        const id = await pridejObjednavku(sheets, p.jmeno, parseFloat((p.kusy||'0').toString().replace(',','.')), p.datum, p.produkt||'vajicka', '');
        result = { status: 'ok', id }; break;
      }

      case 'list': {
        const data = await getOrdersData(sheets);
        const orders = mapOrders(data);
        const produkt = p.produkt || 'vajicka';
        result = orders
          .filter(o => o.produkt === produkt && o.stav === 'čeká')
          .map(o => ({ id: o.id, jmeno: o.jmeno, kusy: o.kusy, datum: o.datum }));
        break;
      }

      case 'done': {
        const o = await najdiObjednavku(sheets, p.id);
        if (!o) { result = { status: 'error', error: 'Objednávka ' + p.id + ' nenalezena' }; break; }
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `${ORDERS_SHEET}!F${o.radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['doručeno']] } });
        result = { status: 'ok' }; break;
      }

      case 'edit': {
        const o = await najdiObjednavku(sheets, p.id);
        if (!o) { result = { status: 'error', error: 'Objednávka ' + p.id + ' nenalezena' }; break; }
        const kusy = parseFloat((p.kusy||'0').toString().replace(',','.')) || 0;
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `${ORDERS_SHEET}!E${o.radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[kusy]] } });
        // Synchronizace do Požadavků – PŘESNĚ přes RequestID
        if (o.requestId !== '') {
          try {
            const reqData = await getPozadavkyData(sheets);
            for (let i = 1; i < reqData.length; i++) {
              const r = reqData[i] || [];
              if ((r[0]||'').toString() !== o.requestId.toString()) continue;
              const prodColIdx = { vajicka: 4, bedynka: 5, sirup: 6 }[o.produkt];
              const v = o.produkt==='vajicka' ? kusy : (parseInt((r[4]||'0').toString().replace(/\s/g,''))||0);
              const b = o.produkt==='bedynka' ? kusy : (parseInt((r[5]||'0').toString().replace(/\s/g,''))||0);
              const s = o.produkt==='sirup'   ? kusy : (parseFloat((r[6]||'0').toString().replace(/\s/g,'').replace(',','.'))||0);
              const jmeno = (r[1]||'').toString();
              const celkem = cenaZa(jmeno,'vajicka',v) + cenaZa(jmeno,'bedynka',b) + cenaZa(jmeno,'sirup',s);
              await sheets.spreadsheets.values.batchUpdate({
                spreadsheetId: SPREADSHEET_ID,
                requestBody: { valueInputOption: 'USER_ENTERED', data: [
                  { range: `Požadavky!${colLetter(prodColIdx+1)}${i+1}`, values: [[kusy]] },
                  { range: `Požadavky!H${i+1}`, values: [[celkem]] }
                ]}
              });
              break;
            }
          } catch(e) { console.log('edit sync selhal: ' + e.message); }
        }
        result = { status: 'ok' }; break;
      }

      case 'delete': {
        const o = await najdiObjednavku(sheets, p.id);
        if (!o) { result = { status: 'error', error: 'Objednávka ' + p.id + ' nenalezena' }; break; }
        const sheetId = await getSheetId(sheets, ORDERS_SHEET);
        await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: o.radek-1, endIndex: o.radek } } }] } });
        result = { status: 'ok' }; break;
      }

      case 'history': {
        const data = await getOrdersData(sheets);
        const orders = mapOrders(data);
        orders.sort((a,b) => {
          const da = parseDate(a.datum), db = parseDate(b.datum);
          return (db?db.getTime():0) - (da?da.getTime():0);
        });
        result = orders.slice(0, 200).map(o => ({ produkt: o.produkt, jmeno: o.jmeno, kusy: o.kusy, datum: o.datum, dorucen: o.stav === 'doručeno' }));
        break;
      }

      case 'init': {
        const [ordersData, reqData, zakData, skladData, aktualita] = await Promise.all([
          getOrdersData(sheets),
          getPozadavkyData(sheets),
          getZakazniciData(sheets),
          getSkladData(sheets),
          getAktualita(sheets)
        ]);
        const orders = mapOrders(ordersData);
        const lists = { vajicka: [], bedynka: [], sirup: [] };
        orders.forEach(o => {
          if (o.stav === 'čeká' && lists[o.produkt]) lists[o.produkt].push({ id: o.id, jmeno: o.jmeno, kusy: o.kusy, datum: o.datum });
        });
        result = {
          customers: zakData.length <= 1 ? [] : zakData.slice(1).map(r => ({ jmeno: r[0], maPin: r[1] !== '' && r[1] != null })),
          lists,
          requests: mapRequests(reqData).filter(r => r.stav !== 'zrušeno'),
          sklad: mapSklad(skladData),
          aktualita
        };
        break;
      }

      case 'cashflow': {
        const mesice = {};
        const gm = d => {
          const x = parseDate(d);
          if (!x || isNaN(x)) return null;
          return x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0');
        };
        const im = m => { if (!mesice[m]) mesice[m] = { trzba:0, naklady:0, vajicka:0, bedynky:0, sirup:0, tV:0, tB:0, tS:0 }; };
        const data = await getOrdersData(sheets);
        const orders = mapOrders(data);
        orders.forEach(o => {
          const m = gm(o.datum); if (!m) return; im(m);
          const t = cenaZa(o.jmeno, o.produkt, o.kusy);
          mesice[m].trzba += t;
          if (o.produkt === 'vajicka') { mesice[m].tV += t; mesice[m].vajicka += o.kusy; }
          else if (o.produkt === 'bedynka') { mesice[m].tB += t; mesice[m].bedynky += o.kusy; }
          else { mesice[m].tS += t; mesice[m].sirup += o.kusy; }
        });
        let grain = [];
        try { const dCF = await getSheetData(sheets,'Cashflow'); grain = findGrainData(dCF); } catch(e){}
        grain.forEach(g=>{
          const mz=gm(g.datum); if(!mz) return; im(mz);
          mesice[mz].naklady += parseFloat((g.cena||'').toString().replace(/[^\d.,-]/g,'').replace(',','.'))||0;
        });
        const nz=['','Leden','Únor','Březen','Duben','Květen','Červen','Červenec','Srpen','Září','Říjen','Listopad','Prosinec'];
        result = Object.keys(mesice).sort().map(m=>{const pts=m.split('-');return{klic:m,nazev:nz[parseInt(pts[1])]+' '+pts[0],trzba:Math.round(mesice[m].trzba),naklady:Math.round(mesice[m].naklady),zisk:Math.round(mesice[m].trzba-mesice[m].naklady),vajicka:mesice[m].vajicka,bedynky:mesice[m].bedynky,sirup:mesice[m].sirup,trzbaVajicka:Math.round(mesice[m].tV),trzbaBedynky:Math.round(mesice[m].tB),trzbaSirup:Math.round(mesice[m].tS)};});
        break;
      }

      case 'grainOrders': {
        let grain = [];
        try { const dCF = await getSheetData(sheets, 'Cashflow'); grain = findGrainData(dCF); } catch(e){}
        result = grain.map(g => ({
          datum: fmtDate(g.datum),
          cena: Math.round(parseFloat((g.cena||'').toString().replace(/[^\d.,-]/g,'').replace(',','.')) || 0),
          kg: parseFloat((g.kg||'').toString().replace(',','.')) || 0
        }));
        break;
      }

      case 'addRequest': {
        const data = await getPozadavkyData(sheets);
        const id = data.length;
        const v=parseInt(p.vajicka)||0, b=parseInt(p.bedynka)||0, s=parseFloat(p.sirup)||0;
        const celkem = cenaZa(p.jmeno,'vajicka',v) + cenaZa(p.jmeno,'bedynka',b) + cenaZa(p.jmeno,'sirup',s);
        await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: 'Požadavky!A:L', valueInputOption: 'USER_ENTERED', requestBody: { values: [[id,p.jmeno,dnesCZ(),p.datum,v,b,s,celkem,'čeká na potvrzení','','','']] } });
        var produkty = [];
        if (v > 0) produkty.push(v + ' ks vajíček');
        if (b > 0) produkty.push(b + ' ks bedýnek');
        if (s > 0) produkty.push(s + ' l sirupu');
        await pushToCustomer(sheets, '__admin__', '🛒 Nová objednávka – ' + p.jmeno, produkty.join(', ') + ' · doručit ' + p.datum + ' · ' + celkem + ' Kč', 'objednavka');
        result = { status: 'ok' }; break;
      }

      case 'listRequests': {
        result = mapRequests(await getPozadavkyData(sheets)).filter(r=>r.stav!=='zrušeno');
        break;
      }

      case 'getMyRequests': {
        result = mapRequests(await getPozadavkyData(sheets)).filter(r=>r.jmeno===p.jmeno);
        break;
      }

      case 'confirmRequest': {
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1];
        if(!row){result={status:'error',error:'Not found row '+radek};break;}
        const reqId = (row[0]||'').toString();
        const pd = fmtDate(row[3]), nd = fmtDate(p.navrzenyDatum);
        const isCounter = nd && nd !== '' && nd !== pd;
        const stavRow = (row[8]||'').toString().trim().toLowerCase();
        if (isCounter) {
          const zprava = (p.zprava || '').toString().substring(0, 200);
          await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:K${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['navržen jiný termín', nd, zprava]] } });
          const pushBody = 'Máš návrh termínu doručení na ' + nd + (zprava ? ' · „' + zprava + '"' : '') + ' – otevři appku a potvrď ho.';
          await pushToCustomer(sheets, row[1], 'Dvůr Pod Dubem', pushBody, 'navrh');
        } else if (stavRow === 'potvrzeno') {
          const datum = nd||pd;
          await pushToCustomer(sheets, row[1], 'Dvůr Pod Dubem', 'Tvoje objednávka je potvrzená na ' + datum + '. Děkujeme!', 'potvrzeno');
        } else {
          const datum = nd||pd;
          // Dvojitá ochrana: existují už řádky s tímto RequestID?
          const existOrders = mapOrders(await getOrdersData(sheets)).filter(o => o.requestId === reqId && reqId !== '');
          if (existOrders.length === 0) {
            const v=parseInt((row[4]||'0').toString().replace(/\s/g,''))||0;
            const b=parseInt((row[5]||'0').toString().replace(/\s/g,''))||0;
            const s=parseFloat((row[6]||'0').toString().replace(/\s/g,'').replace(',','.'))||0;
            if(v>0) await pridejObjednavku(sheets,row[1],v,datum,'vajicka',reqId);
            if(b>0) await pridejObjednavku(sheets,row[1],b,datum,'bedynka',reqId);
            if(s>0) await pridejObjednavku(sheets,row[1],s,datum,'sirup',reqId);
          }
          await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:J${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['potvrzeno','']] } });
          await pushToCustomer(sheets, row[1], 'Dvůr Pod Dubem', 'Tvoje objednávka je potvrzená na ' + datum + '. Děkujeme!', 'potvrzeno');
        }
        result = { status: 'ok' }; break;
      }

      case 'rejectRequest': {
        // Zrušení požadavku + úklid případných už zapsaných objednávek s tímto RequestID
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1];
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['zrušeno']] } });
        if (row) {
          const reqId = (row[0]||'').toString();
          if (reqId !== '') {
            try {
              const orders = mapOrders(await getOrdersData(sheets)).filter(o => o.requestId === reqId && o.stav === 'čeká');
              const sheetId = await getSheetId(sheets, ORDERS_SHEET);
              // mazat od konce, ať se neposunou indexy
              for (const o of orders.sort((a,b)=>b.radek-a.radek)) {
                await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: o.radek-1, endIndex: o.radek } } }] } });
              }
            } catch(e) { console.log('reject úklid selhal: ' + e.message); }
          }
        }
        result = { status: 'ok' }; break;
      }

      case 'acceptSuggested': {
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1];
        if(!row){result={status:'error',error:'Not found row '+radek};break;}
        const nd = fmtDate(row[9]);
        if(!nd){result={status:'error',error:'Žádný navržený termín'};break;}
        const stavRow = (row[8]||'').toString().trim().toLowerCase();
        if (stavRow === 'potvrzeno') { result = { status: 'ok', skipped: 'already-confirmed' }; break; }
        const reqId = (row[0]||'').toString();
        const existOrders = mapOrders(await getOrdersData(sheets)).filter(o => o.requestId === reqId && reqId !== '');
        if (existOrders.length === 0) {
          const v=parseInt((row[4]||'0').toString().replace(/\s/g,''))||0;
          const b=parseInt((row[5]||'0').toString().replace(/\s/g,''))||0;
          const s=parseFloat((row[6]||'0').toString().replace(/\s/g,'').replace(',','.'))||0;
          if(v>0) await pridejObjednavku(sheets,row[1],v,nd,'vajicka',reqId);
          if(b>0) await pridejObjednavku(sheets,row[1],b,nd,'bedynka',reqId);
          if(s>0) await pridejObjednavku(sheets,row[1],s,nd,'sirup',reqId);
        }
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!D${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[nd]] } });
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:K${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['potvrzeno', '', '']] } });
        await pushToCustomer(sheets, '__admin__', '✅ Termín přijat – ' + row[1], row[1] + ' přijal/a termín ' + nd + '. Objednávka je potvrzená.', 'prijato');
        await pushToCustomer(sheets, row[1], 'Dvůr Pod Dubem', 'Tvoje objednávka je potvrzená na ' + nd + '. Děkujeme!', 'potvrzeno');
        result = { status: 'ok' }; break;
      }

      case 'markPaid': {
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1];
        if(!row){result={status:'error',error:'Not found row '+radek};break;}
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!L${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['ano']] } });
        const celkem = (row[7]||'').toString();
        await pushToCustomer(sheets, '__admin__', '💰 Zaplaceno – ' + row[1], row[1] + ' označil/a objednávku (' + celkem + ' Kč, doručení ' + fmtDate(row[3]) + ') jako zaplacenou.', 'platba');
        result = { status: 'ok' }; break;
      }

      case 'getSklad': {
        result = mapSklad(await getSkladData(sheets)); break;
      }
      case 'setSklad': {
        const produkt = (p.produkt||'').toLowerCase();
        const becomingAvailable = p.dostupne==='ano'||p.dostupne==='true'||p.dostupne==='1';
        let wasUnavailable = false;
        try { const prev = mapSklad(await getSkladData(sheets)); wasUnavailable = prev[produkt] && prev[produkt].dostupne===false; } catch(e){}
        await setSklad(sheets, produkt, becomingAvailable, p.od||'', p.popis||'');
        if (becomingAvailable && wasUnavailable && (produkt==='bedynka'||produkt==='sirup')) {
          const nazev = produkt==='bedynka' ? 'Bedýnky' : 'Sirup';
          await pushToCustomer(sheets, null, 'Dvůr Pod Dubem', nazev + ' jsou opět skladem – můžeš objednat!', 'sklad');
        }
        result = { status: 'ok' }; break;
      }

      case 'getAktualita': { result = await getAktualita(sheets); break; }
      case 'setAktualita': {
        await setAktualita(sheets, p.nadpis||'', p.text||'', p.foto||'');
        await pushToCustomer(sheets, null, p.nadpis||'Dvůr Pod Dubem 🌿', (p.text||'').substring(0,120)||'Mrkni na novou aktualitu v appce!', 'aktualita');
        result = { status: 'ok' }; break;
      }
      case 'clearAktualita': { result = { status: 'ok', cleared: await clearAktualita(sheets) }; break; }
      case 'editAktualita': {
        const edited = await editAktualita(sheets, p.nadpis||'', p.text||'', p.foto||'');
        if (!edited) await setAktualita(sheets, p.nadpis||'', p.text||'', p.foto||'');
        result = { status: 'ok' }; break;
      }

      case 'vapidKey': { result = { key: VAPID_PUBLIC }; break; }
      case 'subscribe': {
        await savePushSub(sheets, p.jmeno||'', p.endpoint||'', p.p256dh||'', p.auth||'');
        result = { status: 'ok' }; break;
      }

      case 'getCustomers': {
        const data = await getZakazniciData(sheets);
        result = data.length<=1?[]:data.slice(1).map(r=>({jmeno:r[0],maPin:r[1]!==''&&r[1]!=null}));
        break;
      }
      case 'saveCustomer': {
        const data = await getZakazniciData(sheets);
        let found = false;
        for (let i=1;i<data.length;i++) {
          if(data[i][0]===p.jmeno){await sheets.spreadsheets.values.update({spreadsheetId:SPREADSHEET_ID,range:`Zákazníci!B${i+1}`,valueInputOption:'USER_ENTERED',requestBody:{values:[[p.pin]]}});found=true;break;}
        }
        if(!found){await sheets.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Zákazníci!A:C',valueInputOption:'USER_ENTERED',requestBody:{values:[[p.jmeno,p.pin,dnesCZ()]]}});}
        result = { status: 'ok' }; break;
      }
      case 'checkPin': {
        const data = await getZakazniciData(sheets);
        let ok = false;
        for(let i=1;i<data.length;i++){if(data[i][0]===p.jmeno){ok=data[i][1]?.toString()===p.pin?.toString();break;}}
        result = { status: ok?'ok':'wrong' }; break;
      }

      default: result = { status: 'error', error: 'Unknown action' };
    }
    return { statusCode: 200, headers: hdrs, body: JSON.stringify(result) };
  } catch(err) {
    console.error(err);
    return { statusCode: 500, headers: hdrs, body: JSON.stringify({ status: 'error', error: err.message }) };
  }
};
