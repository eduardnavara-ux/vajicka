// netlify/functions/updateOrder.js
// Admin: potvrzení objednávky, návrh termínu, označení jako doručeno

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
    const { orderId, stav, navrzeny_termin } = JSON.parse(event.body);

    const auth = await getAuth();
    const sheets = google.sheets({ version: 'v4', auth });

    // Najít řádek podle ID (sloupec A)
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:A`,
    });
    const ids = resp.data.values || [];
    const rowIndex = ids.findIndex((r, i) => i > 0 && r[0] == orderId);

    if (rowIndex === -1) {
      return { statusCode: 404, body: JSON.stringify({ error: 'Objednávka nenalezena' }) };
    }

    const sheetRow = rowIndex + 1; // Google Sheets je 1-based

    // Aktualizovat sloupec L (stav) a M (navržený termín)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!L${sheetRow}:M${sheetRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[stav, navrzeny_termin || '']],
      },
    });

    return {
      statusCode: 200,
      body: JSON.stringify({ success: true }),
    };
  } catch (err) {
    console.error('updateOrder error:', err);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
