const { google } = require('googleapis');

const SPREADSHEET_ID = '1ORfd4FhxKsJIuk22WvdoW6sLNpRNBFK39q3yt5YAj8Q';
const DISCOUNT_NAMES = ['Danča', 'Kristýna', 'Renata', 'Andrea', 'Marcela', 'Dáša'];
const SUMA_ROW = 30;
const HEADER_ROW = 2;
const FIRST_COL = 2;
const DATA_START_ROW = 4;
const NTFY_TOPIC = 'dvurpoddubem-objednavky';
// VAPID klíče pro web push – nastav v Netlify env (VAPID_PUBLIC, VAPID_PRIVATE)
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
  // Již ve formátu DD.MM.YYYY
  if (s.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) return s;
  // Excel sériové číslo (číslo bez teček)
  if (s.match(/^\d+$/) && parseInt(s) > 1000) {
    const serial = parseInt(s);
    // Excel epoch: 1. 1. 1900 = 1, ale s bug na den 60 (29.2.1900 neexistuje)
    const date = new Date((serial - 25569) * 86400 * 1000);
    if (!isNaN(date)) return date.getUTCDate() + '.' + (date.getUTCMonth() + 1) + '.' + date.getUTCFullYear();
  }
  // ISO string
  const d = new Date(s);
  if (!isNaN(d)) return d.getUTCDate() + '.' + (d.getUTCMonth() + 1) + '.' + d.getUTCFullYear();
  return s;
}

function parseDate(val) {
  if (!val) return null;
  const s = val.toString().trim();
  // DD.MM.YYYY
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  // Excel sériové číslo
  if (s.match(/^\d+$/) && parseInt(s) > 1000) {
    return new Date((parseInt(s) - 25569) * 86400 * 1000);
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

async function sendNtfy(jmeno, datum, vajicka, bedynka, sirup) {
  try {
    let msg = '';
    if (vajicka > 0) msg += vajicka + ' ks vajíček  ';
    if (bedynka > 0) msg += bedynka + 'x bedýnka  ';
    if (sirup > 0) msg += sirup + 'l sirupu  ';
    msg += '· doručení ' + datum;
    // JSON publish – zvládá diakritiku i emoji; await zajistí dokončení před koncem funkce
    await fetch('https://ntfy.sh/', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        topic: NTFY_TOPIC,
        title: '🛒 ' + jmeno + ' – nová objednávka',
        message: msg,
        priority: 5,
        tags: ['egg']
      })
    });
  } catch (e) { /* notifikace nesmí shodit objednávku */ }
}

async function getSheetData(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:Z`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  return res.data.values || [];
}

async function getSheetColors(sheets, sheetName) {
  const res = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID, ranges: [`${sheetName}!A:Z`],
    fields: 'sheets.data.rowData.values.userEnteredFormat.backgroundColor'
  });
  return res.data.sheets?.[0]?.data?.[0]?.rowData || [];
}

function getCustomers(headers) {
  const c = [];
  for (let i = FIRST_COL - 1; i < headers.length; i++) {
    if (headers[i] && headers[i] !== '') c.push({ jmeno: headers[i], col: i });
    else break;
  }
  return c;
}

// Detekce barev buněk (oranžová #F4A623 = čekající, zelená #93C47D = doručeno)
function isOrange(bg) {
  return !!(bg && bg.red > 0.9 && bg.green > 0.6 && bg.green < 0.72 && bg.blue < 0.2);
}
function isGreen(bg) {
  return !!(bg && bg.red > 0.45 && bg.red < 0.7 && bg.green > 0.7 && bg.green < 0.85 && bg.blue > 0.38 && bg.blue < 0.6);
}

async function getSheetId(sheets, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  return res.data.sheets.find(s => s.properties.title === sheetName)?.properties?.sheetId;
}

async function getOrCreateCustomerCol(sheets, sheetName, jmeno) {
  const data = await getSheetData(sheets, sheetName);
  const headers = data[HEADER_ROW - 1] || [];
  let lastCustomerCol = FIRST_COL - 1;
  for (let i = FIRST_COL - 1; i < headers.length; i++) {
    if (headers[i] === jmeno) return i + 1;
    if (headers[i] && headers[i] !== '') lastCustomerCol = i + 1;
    else break;
  }
  const newCol = lastCustomerCol + 1;
  const sheetId = await getSheetId(sheets, sheetName);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ insertDimension: { range: { sheetId, dimension: 'COLUMNS', startIndex: lastCustomerCol, endIndex: newCol }, inheritFromBefore: false } }] }
  });
  const col = colLetter(newCol);
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: [
      { range: `${sheetName}!${col}${HEADER_ROW}`, values: [[jmeno]] },
      { range: `${sheetName}!${col}${SUMA_ROW}`, values: [[`=SUM(${col}${DATA_START_ROW}:${col}${SUMA_ROW - 1})`]] }
    ]}
  });
  await updatePrehled(sheets, sheetName);
  return newCol;
}

async function updatePrehled(sheets, sheetName) {
  const data = await getSheetData(sheets, sheetName);
  const customers = getCustomers(data[HEADER_ROW - 1] || []);
  if (customers.length === 0) return;
  const fl = colLetter(customers[0].col + 1);
  const ll = colLetter(customers[customers.length - 1].col + 1);
  let trzba;
  if (sheetName === 'Vajíčka') {
    const disc = [], full = [];
    customers.forEach(c => { const r = colLetter(c.col + 1) + SUMA_ROW; if (DISCOUNT_NAMES.includes(c.jmeno)) disc.push(r); else full.push(r); });
    trzba = (disc.length ? `(${disc.join('+')})*9` : '') + (full.length ? (disc.length ? '+' : '') + `(${full.join('+')})*10` : '');
  } else if (sheetName === 'Bedýnky') trzba = `SUM(${fl}${SUMA_ROW}:${ll}${SUMA_ROW})*490`;
  else trzba = `SUM(${fl}${SUMA_ROW}:${ll}${SUMA_ROW})*190`;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: [
      { range: `${sheetName}!L3`, values: [[`=SUM(${fl}${SUMA_ROW}:${ll}${SUMA_ROW})`]] },
      { range: `${sheetName}!L4`, values: [[`=${trzba}`]] }
    ]}
  });
}

async function zpracujObjednavku(sheets, jmeno, kusy, datum, produkt) {
  const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' }[produkt];
  const col = await getOrCreateCustomerCol(sheets, sn, jmeno);
  const data = await getSheetData(sheets, sn);
  // Hledej první prázdný řádek v datové oblasti (DATA_START_ROW až SUMA_ROW-1)
  let targetRow = -1;
  for (let r = DATA_START_ROW; r < SUMA_ROW; r++) {
    const cellVal = data[r - 1]?.[0];
    if (!cellVal || cellVal === '') { targetRow = r; break; }
  }
  if (targetRow === -1) targetRow = SUMA_ROW - 1;

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: 'USER_ENTERED', data: [
      { range: `${sn}!A${targetRow}`, values: [[datum]] },
      { range: `${sn}!${colLetter(col)}${targetRow}`, values: [[parseInt(kusy)]] }
    ]}
  });
  const sheetId = await getSheetId(sheets, sn);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ repeatCell: {
      range: { sheetId, startRowIndex: targetRow - 1, endRowIndex: targetRow, startColumnIndex: col - 1, endColumnIndex: col },
      cell: { userEnteredFormat: { backgroundColor: { red: 0.957, green: 0.651, blue: 0.137 } } },
      fields: 'userEnteredFormat.backgroundColor'
    }}]}
  });
}

// Čekající objednávky (oranžové buňky) jednoho produktu – používá 'list' i 'init'
async function listProduct(sheets, produkt) {
  const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' }[produkt];
  const [data, colorRows] = await Promise.all([getSheetData(sheets, sn), getSheetColors(sheets, sn)]);
  const customers = getCustomers(data[HEADER_ROW - 1] || []);
  const res = [];
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const dc = data[i][0]; if (!dc) continue;
    if (typeof dc === 'string' && isNaN(Date.parse(dc)) && !dc.match(/^\d/)) continue;
    const ds = fmtDate(dc);
    if (!ds) continue;
    for (const c of customers) {
      const val = data[i][c.col]; if (!val || val === '') continue;
      const bg = colorRows[i]?.values?.[c.col]?.userEnteredFormat?.backgroundColor;
      if (isOrange(bg)) res.push({ radek: i + 1, sloupec: c.col + 1, jmeno: c.jmeno, kusy: val, datum: ds });
    }
  }
  return res;
}

// Historie všech objednávek (oranžové i zelené buňky) napříč produkty
async function historyAll(sheets) {
  const produkty = ['vajicka', 'bedynka', 'sirup'];
  const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' };
  const fetched = await Promise.all(produkty.map(p =>
    Promise.all([getSheetData(sheets, sn[p]), getSheetColors(sheets, sn[p])])
  ));
  const out = [];
  produkty.forEach((p, idx) => {
    const [data, colorRows] = fetched[idx];
    const customers = getCustomers(data[HEADER_ROW - 1] || []);
    for (let i = DATA_START_ROW - 1; i < data.length; i++) {
      const dc = data[i][0]; if (!dc) continue;
      if (typeof dc === 'string' && isNaN(Date.parse(dc)) && !dc.match(/^\d/)) continue;
      const ds = fmtDate(dc);
      if (!ds) continue;
      const d = parseDate(dc);
      const ts = d ? d.getTime() : 0;
      for (const c of customers) {
        const val = data[i][c.col]; if (!val || val === '') continue;
        const bg = colorRows[i]?.values?.[c.col]?.userEnteredFormat?.backgroundColor;
        const orange = isOrange(bg), green = isGreen(bg);
        if (!orange && !green) continue;
        out.push({ produkt: p, jmeno: c.jmeno, kusy: val, datum: ds, dorucen: green, ts });
      }
    }
  });
  out.sort((a, b) => b.ts - a.ts);
  return out.slice(0, 200).map(o => ({ produkt: o.produkt, jmeno: o.jmeno, kusy: o.kusy, datum: o.datum, dorucen: o.dorucen }));
}

function mapRequests(data) {
  return data.length <= 1 ? [] : data.slice(1).map((r, i) => ({
    radek: i + 2, id: r[0], jmeno: r[1], datumPozadavku: fmtDate(r[2]), datumDoruceni: fmtDate(r[3]),
    vajicka: r[4] || 0, bedynka: r[5] || 0, sirup: r[6] || 0, celkem: r[7] || 0,
    stav: r[8] || 'čeká na potvrzení', navrzenyTermin: fmtDate(r[9])
  }));
}

async function getPozadavkyData(sheets) {
  try { return await getSheetData(sheets, 'Požadavky'); }
  catch {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Požadavky' } } }] } });
    const hdr = [['ID','Zákazník','Datum požadavku','Datum doručení','Vajíčka','Bedýnka','Sirup','Celkem','Stav','Navržený termín']];
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: 'Požadavky!A1:J1', valueInputOption: 'USER_ENTERED', requestBody: { values: hdr } });
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

// ── SKLAD (dostupnost produktů) ──
// Struktura listu Sklad:
// A: produkt (sirup|bedynka|vajicka), B: dostupne (ano|ne), C: od (datum), D: popis
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
  // prázdný nebo bez hlavičky → naplň výchozími hodnotami
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

// ════════════════════════════════════════════════════
// WEB PUSH (nativní, bez knihovny web-push)
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

// VAPID JWT (ES256) pro autorizaci u push služby
function vapidJWT(audience) {
  const header = b64urlEncode(JSON.stringify({ typ: 'JWT', alg: 'ES256' }));
  const payload = b64urlEncode(JSON.stringify({
    aud: audience,
    exp: Math.floor(Date.now()/1000) + 12*3600,
    sub: VAPID_SUBJECT
  }));
  const unsigned = header + '.' + payload;
  // privátní klíč (32B raw) → PKCS8 pro Node sign
  const privBuf = b64urlDecode(VAPID_PRIVATE);
  const pubBuf = b64urlDecode(VAPID_PUBLIC); // 65B uncompressed
  const jwk = {
    kty: 'EC', crv: 'P-256',
    d: b64urlEncode(privBuf),
    x: b64urlEncode(pubBuf.slice(1, 33)),
    y: b64urlEncode(pubBuf.slice(33, 65))
  };
  const keyObj = crypto.createPrivateKey({ key: jwk, format: 'jwk' });
  const der = crypto.sign('sha256', Buffer.from(unsigned), { key: keyObj, dsaEncoding: 'ieee-p1363' });
  return unsigned + '.' + b64urlEncode(der);
}

// HKDF (SHA-256)
function hkdf(salt, ikm, info, length) {
  const prk = crypto.createHmac('sha256', salt).update(ikm).digest();
  const out = crypto.createHmac('sha256', prk).update(Buffer.concat([info, Buffer.from([1])])).digest();
  return out.slice(0, length);
}

// Zašifruje payload dle RFC 8291 (aes128gcm)
function encryptPayload(payloadStr, p256dhB64, authB64) {
  const clientPub = b64urlDecode(p256dhB64);     // 65B
  const auth = b64urlDecode(authB64);            // 16B
  const salt = crypto.randomBytes(16);

  const ecdh = crypto.createECDH('prime256v1');
  ecdh.generateKeys();
  const serverPub = ecdh.getPublicKey();         // 65B
  const sharedSecret = ecdh.computeSecret(clientPub);

  // PRK_key + IKM dle RFC 8291
  const authInfo = Buffer.concat([Buffer.from('WebPush: info\0'), clientPub, serverPub]);
  const ikm = hkdf(auth, sharedSecret, authInfo, 32);

  const keyInfo = Buffer.concat([Buffer.from('Content-Encoding: aes128gcm\0')]);
  const cek = hkdf(salt, ikm, keyInfo, 16);
  const nonceInfo = Buffer.concat([Buffer.from('Content-Encoding: nonce\0')]);
  const nonce = hkdf(salt, ikm, nonceInfo, 12);

  const payload = Buffer.from(payloadStr, 'utf8');
  const padded = Buffer.concat([payload, Buffer.from([0x02])]); // delimiter, no padding
  const cipher = crypto.createCipheriv('aes-128-gcm', cek, nonce);
  const encrypted = Buffer.concat([cipher.update(padded), cipher.final(), cipher.getAuthTag()]);

  // aes128gcm header: salt(16) | rs(4=4096) | idlen(1=65) | serverPub(65)
  const header = Buffer.concat([
    salt,
    Buffer.from([0,0,0x10,0x00]),
    Buffer.from([serverPub.length]),
    serverPub
  ]);
  return Buffer.concat([header, encrypted]);
}

function sendOnePush(sub, payloadStr) {
  return new Promise((resolve) => {
    try {
      if (!VAPID_PUBLIC || !VAPID_PRIVATE) return resolve({ ok:false, gone:false });
      const url = new URL(sub.endpoint);
      const audience = url.origin;
      const body = encryptPayload(payloadStr, sub.p256dh, sub.auth);
      const jwt = vapidJWT(audience);
      const https = require('https');
      const req = https.request({
        hostname: url.hostname,
        path: url.pathname + url.search,
        method: 'POST',
        headers: {
          'Content-Type': 'application/octet-stream',
          'Content-Encoding': 'aes128gcm',
          'Content-Length': body.length,
          'TTL': '86400',
          'Urgency': 'high',
          'Authorization': 'vapid t=' + jwt + ', k=' + VAPID_PUBLIC
        }
      }, (res) => {
        // 404/410 = subscription zaniklo → smazat
        const gone = res.statusCode === 404 || res.statusCode === 410;
        res.on('data', ()=>{}); res.on('end', ()=> resolve({ ok: res.statusCode>=200&&res.statusCode<300, gone }));
      });
      req.on('error', () => resolve({ ok:false, gone:false }));
      req.write(body); req.end();
    } catch(e) { resolve({ ok:false, gone:false }); }
  });
}

// ── Úložiště odběratelů: list Push ──
// A: jmeno, B: endpoint, C: p256dh, D: auth
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
  // duplikát endpointu? aktualizuj jméno; jinak přidej
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
  if (!VAPID_PUBLIC) return;
  const subs = await getSubsFor(sheets, jmeno);
  const payload = JSON.stringify({ title, body, tag: tag||'dpd', url: '/zakaznik' });
  const goneRows = [];
  for (const s of subs) {
    const r = await sendOnePush(s, payload);
    if (r.gone) goneRows.push(s.row);
  }
  // smaž zaniklé (od konce, ať nesedne indexace)
  for (const row of goneRows.sort((a,b)=>b-a)) {
    try { await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: `Push!A${row}:D${row}` }); } catch(e){}
  }
}

const hdrs = { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' };

exports.handler = async function(event) {
  const p = event.queryStringParameters || {};
  try {
    const sheets = await getSheets();
    let result;
    switch (p.action) {
      case 'add': {
        await zpracujObjednavku(sheets, p.jmeno, p.kusy, p.datum, p.produkt||'vajicka');
        result = { status: 'ok' }; break;
      }
      case 'done': {
        const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' }[p.produkt||'vajicka'];
        const sheetId = await getSheetId(sheets, sn);
        const row = parseInt(p.radek), col = parseInt(p.sloupec);
        await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ repeatCell: {
          range: { sheetId, startRowIndex: row-1, endRowIndex: row, startColumnIndex: col-1, endColumnIndex: col },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.576, green: 0.769, blue: 0.49 } } },
          fields: 'userEnteredFormat.backgroundColor'
        }}]}});
        result = { status: 'ok' }; break;
      }
      case 'edit': {
        const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' }[p.produkt||'vajicka'];
        const row = parseInt(p.radek), col = parseInt(p.sloupec);
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `${sn}!${colLetter(col)}${row}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[parseInt(p.kusy)]] } });
        const sheetId = await getSheetId(sheets, sn);
        await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ repeatCell: {
          range: { sheetId, startRowIndex: row-1, endRowIndex: row, startColumnIndex: col-1, endColumnIndex: col },
          cell: { userEnteredFormat: { backgroundColor: { red: 0.957, green: 0.651, blue: 0.137 } } },
          fields: 'userEnteredFormat.backgroundColor'
        }}]}});
        result = { status: 'ok' }; break;
      }
      case 'delete': {
        const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' }[p.produkt||'vajicka'];
        const row = parseInt(p.radek), col = parseInt(p.sloupec);
        await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: `${sn}!${colLetter(col)}${row}` });
        const sheetId = await getSheetId(sheets, sn);
        await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ repeatCell: {
          range: { sheetId, startRowIndex: row-1, endRowIndex: row, startColumnIndex: col-1, endColumnIndex: col },
          cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
          fields: 'userEnteredFormat.backgroundColor'
        }}]}});
        const data = await getSheetData(sheets, sn);
        const rd = data[row-1] || [];
        if (rd.slice(1).every(v => !v || v === '')) {
          await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: row-1, endIndex: row } } }] } });
        }
        result = { status: 'ok' }; break;
      }
      case 'list': {
        result = await listProduct(sheets, p.produkt||'vajicka'); break;
      }
      case 'init': {
        // Vše pro start appky v jednom volání
        const [vajicka, bedynka, sirup, reqData, zakData, skladData] = await Promise.all([
          listProduct(sheets, 'vajicka'),
          listProduct(sheets, 'bedynka'),
          listProduct(sheets, 'sirup'),
          getPozadavkyData(sheets),
          getZakazniciData(sheets),
          getSkladData(sheets)
        ]);
        result = {
          customers: zakData.length <= 1 ? [] : zakData.slice(1).map(r => ({ jmeno: r[0], maPin: r[1] !== '' && r[1] != null })),
          lists: { vajicka, bedynka, sirup },
          requests: mapRequests(reqData).filter(r => r.stav !== 'zrušeno'),
          sklad: mapSklad(skladData)
        };
        break;
      }
      case 'getSklad': {
        result = mapSklad(await getSkladData(sheets)); break;
      }
      case 'setSklad': {
        const produkt = (p.produkt||'').toLowerCase();
        const becomingAvailable = p.dostupne==='ano'||p.dostupne==='true'||p.dostupne==='1';
        // zjisti předchozí stav, ať pushneme jen při přechodu nedostupné→dostupné
        let wasUnavailable = false;
        try { const prev = mapSklad(await getSkladData(sheets)); wasUnavailable = prev[produkt] && prev[produkt].dostupne===false; } catch(e){}
        await setSklad(sheets, produkt, becomingAvailable, p.od||'', p.popis||'');
        if (becomingAvailable && wasUnavailable && (produkt==='bedynka'||produkt==='sirup')) {
          const nazev = produkt==='bedynka' ? 'Bedýnky' : 'Sirup';
          await pushToCustomer(sheets, null, 'Dvůr Pod Dubem', nazev + ' jsou opět skladem – můžeš objednat!', 'sklad');
        }
        result = { status: 'ok' }; break;
      }
      case 'vapidKey': {
        result = { key: VAPID_PUBLIC }; break;
      }
      case 'subscribe': {
        await savePushSub(sheets, p.jmeno||'', p.endpoint||'', p.p256dh||'', p.auth||'');
        result = { status: 'ok' }; break;
      }
      case 'history': {
        result = await historyAll(sheets); break;
      }
      case 'cashflow': {
        const mesice = {};
        const gm = d => {
          if (!d) return null;
          const s = d.toString().trim();
          // Přeskoč textové záhlavíčky jako "Květen 2026"
          if (!s.match(/^\d/)) return null;
          const x = parseDate(s);
          if (!x || isNaN(x)) return null;
          return x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0');
        };
        const im = m => { if (!mesice[m]) mesice[m] = { trzba:0, naklady:0, vajicka:0, bedynky:0, sirup:0, tV:0, tB:0, tS:0 }; };
        const dV = await getSheetData(sheets,'Vajíčka'); const zV = getCustomers(dV[HEADER_ROW-1]||[]);
        for (let i=DATA_START_ROW-1;i<dV.length;i++) {
          const d=dV[i][0]; if(!d) continue;
          const m=gm(d); if(!m) continue; im(m);
          zV.forEach(z=>{const v=parseInt((dV[i][z.col]||'').toString().replace(/\s/g,''))||0;if(v>0){const t=v*(DISCOUNT_NAMES.includes(z.jmeno)?9:10);mesice[m].trzba+=t;mesice[m].tV+=t;mesice[m].vajicka+=v;}});
          const dz=dV[i][13],cz=dV[i][14]; if(dz&&cz){const mz=gm(dz);if(mz){im(mz);mesice[mz].naklady+=parseFloat((cz||'').toString().replace(/\s/g,'').replace(',','.'))||0;}}
        }
        const dB = await getSheetData(sheets,'Bedýnky'); const zB = getCustomers(dB[HEADER_ROW-1]||[]);
        for (let i=DATA_START_ROW-1;i<dB.length;i++) {
          const d=dB[i][0]; if(!d) continue;
          const m=gm(d); if(!m) continue; im(m);
          zB.forEach(z=>{const v=parseInt((dB[i][z.col]||'').toString().replace(/\s/g,''))||0;if(v>0){mesice[m].trzba+=v*490;mesice[m].tB+=v*490;mesice[m].bedynky+=v;}});
        }
        const dS = await getSheetData(sheets,'Sirup'); const zS = getCustomers(dS[HEADER_ROW-1]||[]);
        for (let i=DATA_START_ROW-1;i<dS.length;i++) {
          const d=dS[i][0]; if(!d) continue;
          const m=gm(d); if(!m) continue; im(m);
          zS.forEach(z=>{const v=parseInt((dS[i][z.col]||'').toString().replace(/\s/g,''))||0;if(v>0){mesice[m].trzba+=v*190;mesice[m].tS+=v*190;mesice[m].sirup+=v;}});
        }
        const nz=['','Leden','Únor','Březen','Duben','Květen','Červen','Červenec','Srpen','Září','Říjen','Listopad','Prosinec'];
        result = Object.keys(mesice).sort().map(m=>{const pts=m.split('-');return{klic:m,nazev:nz[parseInt(pts[1])]+' '+pts[0],trzba:Math.round(mesice[m].trzba),naklady:Math.round(mesice[m].naklady),zisk:Math.round(mesice[m].trzba-mesice[m].naklady),vajicka:mesice[m].vajicka,bedynky:mesice[m].bedynky,sirup:mesice[m].sirup,trzbaVajicka:Math.round(mesice[m].tV),trzbaBedynky:Math.round(mesice[m].tB),trzbaSirup:Math.round(mesice[m].tS)};});
        break;
      }
      case 'addRequest': {
        const data = await getPozadavkyData(sheets);
        const id = data.length;
        const v=parseInt(p.vajicka)||0, b=parseInt(p.bedynka)||0, s=parseFloat(p.sirup)||0;
        const ep = DISCOUNT_NAMES.includes(p.jmeno)?9:10;
        const celkem = v*ep+b*490+Math.round(s*190);
        const dn = new Date(); const ds = dn.getDate()+'.'+(dn.getMonth()+1)+'.'+dn.getFullYear();
        await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: 'Požadavky!A:J', valueInputOption: 'USER_ENTERED', requestBody: { values: [[id,p.jmeno,ds,p.datum,v,b,s,celkem,'čeká na potvrzení','']] } });
        await sendNtfy(p.jmeno, p.datum, v, b, s);
        result = { status: 'ok' }; break;
      }
      case 'listRequests': {
        const data = await getPozadavkyData(sheets);
        result = mapRequests(data).filter(r=>r.stav!=='zrušeno');
        break;
      }
      case 'getMyRequests': {
        const data = await getPozadavkyData(sheets);
        result = mapRequests(data).filter(r=>r.jmeno===p.jmeno);
        break;
      }
      case 'confirmRequest': {
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1];
        if(!row){result={status:'error',error:'Not found row '+radek+' in '+data.length};break;}
        const pd = fmtDate(row[3]), nd = fmtDate(p.navrzenyDatum);
        const isCounter = nd && nd !== '' && nd !== pd;
        if (isCounter) {
          await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:J${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['navržen jiný termín', nd]] } });
          await pushToCustomer(sheets, row[1], 'Dvůr Pod Dubem', 'Máš návrh termínu doručení na ' + nd + ' – otevři appku a potvrď ho.', 'navrh');
        } else {
          const datum = nd||pd;
          const v=parseInt((row[4]||'0').toString().replace(/\s/g,''))||0;
          const b=parseInt((row[5]||'0').toString().replace(/\s/g,''))||0;
          const s=parseFloat((row[6]||'0').toString().replace(/\s/g,'').replace(',','.'))||0;
          console.log('confirmRequest: jmeno='+row[1]+' v='+v+' b='+b+' s='+s+' datum='+datum);
          if(v>0) await zpracujObjednavku(sheets,row[1],v,datum,'vajicka');
          if(b>0) await zpracujObjednavku(sheets,row[1],b,datum,'bedynka');
          if(s>0) await zpracujObjednavku(sheets,row[1],s,datum,'sirup');
          await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:J${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['potvrzeno','']] } });
          await pushToCustomer(sheets, row[1], 'Dvůr Pod Dubem', 'Tvoje objednávka je potvrzená na ' + datum + '. Děkujeme!', 'potvrzeno');
        }
        result = { status: 'ok' }; break;
      }
      case 'rejectRequest': {
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${parseInt(p.radek)}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['zrušeno']] } });
        result = { status: 'ok' }; break;
      }
      case 'acceptSuggested': {
        // Zákazník (nebo admin) přijal navržený termín – zapiš objednávku s datem ze sloupce J
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1];
        if(!row){result={status:'error',error:'Not found row '+radek};break;}
        const nd = fmtDate(row[9]);
        if(!nd){result={status:'error',error:'Žádný navržený termín'};break;}
        const v=parseInt((row[4]||'0').toString().replace(/\s/g,''))||0;
        const b=parseInt((row[5]||'0').toString().replace(/\s/g,''))||0;
        const s=parseFloat((row[6]||'0').toString().replace(/\s/g,'').replace(',','.'))||0;
        if(v>0) await zpracujObjednavku(sheets,row[1],v,nd,'vajicka');
        if(b>0) await zpracujObjednavku(sheets,row[1],b,nd,'bedynka');
        if(s>0) await zpracujObjednavku(sheets,row[1],s,nd,'sirup');
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!D${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [[nd]] } });
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['potvrzeno']] } });
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
        if(!found){const dn=new Date();const ds=dn.getDate()+'.'+(dn.getMonth()+1)+'.'+dn.getFullYear();await sheets.spreadsheets.values.append({spreadsheetId:SPREADSHEET_ID,range:'Zákazníci!A:C',valueInputOption:'USER_ENTERED',requestBody:{values:[[p.jmeno,p.pin,ds]]}});}
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
