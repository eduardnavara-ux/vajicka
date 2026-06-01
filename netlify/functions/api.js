const https = require('https');
const http = require('http');

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwqBirb-5vMuIpIpuy0MWIBBc-ZlwmLlfdHXIa-SqjtzAQFLZfqf26xyvBvmSqpT8_Ung/exec';

function fetchWithRedirects(url, redirectCount, resolve) {
  if (redirectCount > 10) {
    resolve({ statusCode: 500, headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ status: 'error', error: 'Too many redirects' }) });
    return;
  }
  const lib = url.startsWith('https') ? https : http;
  lib.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, (res) => {
    if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
      fetchWithRedirects(res.headers.location, redirectCount + 1, resolve);
      return;
    }
    let data = '';
    res.on('data', chunk => data += chunk);
    res.on('end', () => {
      resolve({
        statusCode: 200,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        body: data
      });
    });
  }).on('error', (e) => {
    resolve({ statusCode: 500, headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ status: 'error', error: e.message }) });
  });
}

exports.handler = async function(event) {
  const params = event.queryStringParameters || {};
  const qs = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
  const url = SCRIPT_URL + '?' + qs;
  return new Promise((resolve) => fetchWithRedirects(url, 0, resolve));
};
