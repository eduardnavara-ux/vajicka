// netlify/functions/getAllOrders.js
// Admin: načte všechny objednávky (nebo filtrované podle stavu)

const { google } = require('googleapis');

const SPREADSHEET_ID = '1ORfd4FhxKsJIuk22WvdoW6sLNpRNBFK39q3yt5YAj8Q';
const SHEET_NAME = 'Objednávky';

async function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  return auth;
}

exports.handler = async (event) => {
  const stavFilter = event.queryStringParameters?.stav || null;

  try {
    const auth = await getAuth();
    const sheets = google.sheets({ version: 'v4', auth });

    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:M`,
    });

    const rows = resp.data.values || [];
    if (rows.length <= 1) {
      return { statusCode: 200, body: JSON.stringify({ orders: [] }) };
    }

    let orders = rows.slice(1).map(row => ({
      id: row[0] || '',
      zakaznik: row[1] || '',
      datum_objednavky: row[2] || '',
      datum_doruceni: row[3] || '',
      vajicka: parseInt(row[4]) || 0,
      bedynka: parseInt(row[5]) || 0,
      sirup: parseFloat(row[6]) || 0,
      cena_vajicek: parseInt(row[7]) || 0,
      cena_bedynek: parseInt(row[8]) || 0,
      cena_sirupu: parseInt(row[9]) || 0,
      celkem: parseInt(row[10]) || 0,
      stav: row[11] || 'čeká na potvrzení',
      navrzeny_termin: row[12] || '',
    }));

    if (stavFilter) {
      orders = orders.filter(o => o.stav === stavFilter);
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ orders }),
    };
  } catch (err) {
    console.error('getAllOrders error:', err);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
