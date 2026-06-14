// ============================================================================
// Dvůr Pod Dubem – netlify/functions/upload.js
// Přijme fotku (base64 data URL) z admin appky, pošle na Cloudinary, vrátí URL.
//
// KAM: netlify/functions/upload.js  (vedle api.js)
// CLOUDINARY: cloud name = dxtlvckbt, unsigned preset = ml_default
//
// Volá se z admin appky: POST /.netlify/functions/upload  body: { image: "data:image/jpeg;base64,..." }
// Vrací: { url: "https://res.cloudinary.com/..." }  nebo { error: "..." }
//
// NEOTESTOVÁNO automaticky (chat bez Node.js). Po nasazení vyzkoušej nahrání fotky.
// ============================================================================

const CLOUD_NAME = 'dxtlvckbt';
const UPLOAD_PRESET = 'ml_default';

exports.handler = async function (event) {
  const cors = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: cors, body: '' };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: cors, body: JSON.stringify({ error: 'Použij POST' }) };
  }

  try {
    const body = JSON.parse(event.body || '{}');
    const image = body.image; // očekává data URL: "data:image/jpeg;base64,..."
    if (!image || typeof image !== 'string' || image.indexOf('data:image/') !== 0) {
      return { statusCode: 400, headers: cors, body: JSON.stringify({ error: 'Chybí obrázek nebo špatný formát' }) };
    }

    // Cloudinary unsigned upload – pošleme data URL přímo do pole "file"
    const form = new URLSearchParams();
    form.append('file', image);
    form.append('upload_preset', UPLOAD_PRESET);

    const resp = await fetch(`https://api.cloudinary.com/v1_1/${CLOUD_NAME}/image/upload`, {
      method: 'POST',
      body: form
    });

    const data = await resp.json();
    if (!resp.ok || !data.secure_url) {
      return { statusCode: 502, headers: cors, body: JSON.stringify({ error: (data.error && data.error.message) || 'Upload na Cloudinary selhal' }) };
    }

    return { statusCode: 200, headers: cors, body: JSON.stringify({ url: data.secure_url }) };
  } catch (err) {
    return { statusCode: 500, headers: cors, body: JSON.stringify({ error: err.message }) };
  }
};
