const https = require('https');

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwyhnFVcAYfvjs_Lp0iRLsNf1HBJlv6KpArkSbHDZ6GxglOLtHnx2D2x44cPaX-c9zTZg/exec';

exports.handler = async function(event) {
  const params = event.queryStringParameters || {};
  const qs = Object.keys(params).map(k => k+'='+encodeURIComponent(params[k])).join('&');
  const url = SCRIPT_URL + '?' + qs;

  return new Promise((resolve) => {
    https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        resolve({
          statusCode: 200,
          headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
          },
          body: data
        });
      });
    }).on('error', (e) => {
      resolve({
        statusCode: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        body: JSON.stringify({ status: 'error', error: e.message })
      });
    });
  });
};
