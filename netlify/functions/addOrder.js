// netlify/functions/addOrder.js
// Zapisuje zákaznickou objednávku do Google Sheets – list "Objednávky"

const { google } = require('googleapis');

const SPREADSHEET_ID = '1ORfd4FhxKsJIuk22WvdoW6sLNpRNBFK39q3yt5YAj8Q';
const SHEET_NAME = 'Objednávky';

async function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    const order = JSON.parse(event.body);
    const auth = await getAuth();
    const sheets = google.sheets({ version: 'v4', auth });

    // Přidáme řádek do listu "Objednávky"
    // Sloupce: A=ID, B=Zákazník, C=Datum objednávky, D=Datum doručení,
    //          E=Vajíčka (ks), F=Bedýnka (ks), G=Sirup (l),
    //          H=Cena vajíček, I=Cena bedýnek, J=Cena sirupu,
    //          K=Celkem, L=Stav, M=Navržený termín

    // Nejdřív zjistíme počet řádků pro ID
    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:A`,
    });
    const rows = existing.data.values || [];
    const newId = rows.length; // první řádek = hlavička, ID = počet dat

    const rowData = [
      newId,
      order.zakaznik,
      order.datum_objednavky,
      order.datum_doruceni,
      order.vajicka || 0,
      order.bedynka || 0,
      order.sirup || 0,
      order.cena_vajicek || 0,
      order.cena_bedynek || 0,
      order.cena_sirupu || 0,
      order.celkem || 0,
      'čeká na potvrzení',
      '', // navržený termín – prázdné
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:M`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [rowData] },
    });

    return {
      statusCode: 200,
      body: JSON.stringify({ success: true, id: newId }),
    };
  } catch (err) {
    console.error('addOrder error:', err);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
