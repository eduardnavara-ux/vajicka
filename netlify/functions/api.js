const { google } = require('googleapis');

const SPREADSHEET_ID = '1ORfd4FhxKsJIuk22WvdoW6sLNpRNBFK39q3yt5YAj8Q';
const DISCOUNT_NAMES = ['Danča', 'Kristýna', 'Renata', 'Andrea', 'Marcela', 'Dáša'];
const SUMA_ROW = 30;
const HEADER_ROW = 2;
const FIRST_COL = 2;
const DATA_START_ROW = 4;
const NTFY_TOPIC = 'dvurpoddubem-objednavky';

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

function sendNtfy(jmeno, datum, vajicka, bedynka, sirup) {
  const https = require('https');
  let msg = '';
  if (vajicka > 0) msg += vajicka + ' ks vajíček  ';
  if (bedynka > 0) msg += bedynka + 'x bedýnka  ';
  if (sirup > 0) msg += sirup + 'l sirupu  ';
  msg += '· doručení ' + datum;
  const body = Buffer.from(msg);
  const req = https.request({
    hostname: 'ntfy.sh', port: 443, path: '/' + NTFY_TOPIC, method: 'POST',
    headers: { 'Title': '🛒 ' + jmeno + ' – nová objednávka', 'Priority': 'high', 'Tags': 'egg', 'Content-Length': body.length }
  });
  req.on('error', () => {}); req.write(body); req.end();
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
  let targetRow = data.length + 1;
  for (let r = DATA_START_ROW; r <= data.length + 1; r++) {
    if (!data[r - 1]?.[0]) { targetRow = r; break; }
  }
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

const hdrs = { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' };

exports.handler = async function(event) {
  const p = event.queryStringParameters || {};
  if (p.action === 'addRequest') sendNtfy(p.jmeno, p.datum, parseInt(p.vajicka)||0, parseInt(p.bedynka)||0, parseFloat(p.sirup)||0);
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
        const sn = { vajicka: 'Vajíčka', bedynka: 'Bedýnky', sirup: 'Sirup' }[p.produkt||'vajicka'];
        const data = await getSheetData(sheets, sn);
        const colorRows = await getSheetColors(sheets, sn);
        const customers = getCustomers(data[HEADER_ROW-1]||[]);
        const res = [];
        for (let i = DATA_START_ROW-1; i < data.length; i++) {
          const dc = data[i][0]; if (!dc) continue;
          if (typeof dc === 'string' && isNaN(Date.parse(dc)) && !dc.match(/^\d/)) continue;
          const ds = fmtDate(dc);
          if (!ds) continue;
          for (const c of customers) {
            const val = data[i][c.col]; if (!val || val === '') continue;
            const bg = colorRows[i]?.values?.[c.col]?.userEnteredFormat?.backgroundColor;
            const isOrange = bg && bg.red > 0.9 && bg.green > 0.6 && bg.green < 0.72 && bg.blue < 0.2;
            if (isOrange) res.push({ radek: i+1, sloupec: c.col+1, jmeno: c.jmeno, kusy: val, datum: ds });
          }
        }
        result = res; break;
      }
      case 'cashflow': {
        const mesice = {};
        const gm = d => { const x = parseDate(d); if (!x || isNaN(x)) return null; return x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0'); };
        const im = m => { if (!mesice[m]) mesice[m] = { trzba:0, naklady:0, vajicka:0, bedynky:0, sirup:0 }; };
        const dV = await getSheetData(sheets,'Vajíčka'); const zV = getCustomers(dV[HEADER_ROW-1]||[]);
        for (let i=DATA_START_ROW-1;i<dV.length;i++) {
          const d=dV[i][0]; if(!d) continue; if(typeof d==='string'&&isNaN(Date.parse(d))&&!d.match(/^\d/)) continue;
          const m=gm(d); if(!m) continue; im(m);
          zV.forEach(z=>{const v=parseInt(dV[i][z.col])||0;if(v>0){mesice[m].trzba+=v*(DISCOUNT_NAMES.includes(z.jmeno)?9:10);mesice[m].vajicka+=v;}});
          const dz=dV[i][13],cz=dV[i][14]; if(dz&&cz){const mz=gm(dz);if(mz){im(mz);mesice[mz].naklady+=parseFloat(cz)||0;}}
        }
        const dB = await getSheetData(sheets,'Bedýnky'); const zB = getCustomers(dB[HEADER_ROW-1]||[]);
        for (let i=DATA_START_ROW-1;i<dB.length;i++) {
          const d=dB[i][0]; if(!d) continue; if(typeof d==='string'&&isNaN(Date.parse(d))&&!d.match(/^\d/)) continue;
          const m=gm(d); if(!m) continue; im(m);
          zB.forEach(z=>{const v=parseInt(dB[i][z.col])||0;if(v>0){mesice[m].trzba+=v*490;mesice[m].bedynky+=v;}});
        }
        const dS = await getSheetData(sheets,'Sirup'); const zS = getCustomers(dS[HEADER_ROW-1]||[]);
        for (let i=DATA_START_ROW-1;i<dS.length;i++) {
          const d=dS[i][0]; if(!d) continue; if(typeof d==='string'&&isNaN(Date.parse(d))&&!d.match(/^\d/)) continue;
          const m=gm(d); if(!m) continue; im(m);
          zS.forEach(z=>{const v=parseInt(dS[i][z.col])||0;if(v>0){mesice[m].trzba+=v*190;mesice[m].sirup+=v;}});
        }
        const nz=['','Leden','Únor','Březen','Duben','Květen','Červen','Červenec','Srpen','Září','Říjen','Listopad','Prosinec'];
        result = Object.keys(mesice).sort().map(m=>{const pts=m.split('-');return{klic:m,nazev:nz[parseInt(pts[1])]+' '+pts[0],trzba:Math.round(mesice[m].trzba),naklady:Math.round(mesice[m].naklady),zisk:Math.round(mesice[m].trzba-mesice[m].naklady),vajicka:mesice[m].vajicka,bedynky:mesice[m].bedynky,sirup:mesice[m].sirup};});
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
        result = { status: 'ok' }; break;
      }
      case 'listRequests': {
        const data = await getPozadavkyData(sheets);
        result = data.length<=1?[]:data.slice(1).map((r,i)=>({radek:i+2,id:r[0],jmeno:r[1],datumPozadavku:fmtDate(r[2]),datumDoruceni:fmtDate(r[3]),vajicka:r[4]||0,bedynka:r[5]||0,sirup:r[6]||0,celkem:r[7]||0,stav:r[8]||'čeká na potvrzení',navrzenyTermin:fmtDate(r[9])})).filter(r=>r.stav!=='zrušeno');
        break;
      }
      case 'getMyRequests': {
        const data = await getPozadavkyData(sheets);
        result = data.length<=1?[]:data.slice(1).map((r,i)=>({radek:i+2,id:r[0],jmeno:r[1],datumPozadavku:fmtDate(r[2]),datumDoruceni:fmtDate(r[3]),vajicka:r[4]||0,bedynka:r[5]||0,sirup:r[6]||0,celkem:r[7]||0,stav:r[8]||'čeká na potvrzení',navrzenyTermin:fmtDate(r[9])})).filter(r=>r.jmeno===p.jmeno);
        break;
      }
      case 'confirmRequest': {
        const data = await getPozadavkyData(sheets);
        const radek = parseInt(p.radek);
        const row = data[radek-1]; if(!row){result={status:'error',error:'Not found'};break;}
        const pd = fmtDate(row[3]), nd = fmtDate(p.navrzenyDatum);
        const isCounter = nd && nd !== '' && nd !== pd;
        if (isCounter) {
          await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:J${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['navržen jiný termín', nd]] } });
        } else {
          const datum = nd||pd;
          const v=parseInt(row[4])||0, b=parseInt(row[5])||0, s=parseFloat(row[6])||0;
          if(v>0) await zpracujObjednavku(sheets,row[1],v,datum,'vajicka');
          if(b>0) await zpracujObjednavku(sheets,row[1],b,datum,'bedynka');
          if(s>0) await zpracujObjednavku(sheets,row[1],s,datum,'sirup');
          await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${radek}:J${radek}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['potvrzeno','']] } });
        }
        result = { status: 'ok' }; break;
      }
      case 'rejectRequest': {
        await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: `Požadavky!I${parseInt(p.radek)}`, valueInputOption: 'USER_ENTERED', requestBody: { values: [['zrušeno']] } });
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
