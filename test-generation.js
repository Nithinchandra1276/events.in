#!/usr/bin/env node
const http = require('http');

const data = JSON.stringify({
  prompt: 'Front view professional stage 20ft wide with neon LED',
  width: 20,
  height: 10,
  ledPattern: 'neon',
  soundSystem: 'stereo',
  equipment: 'both'
});

console.log('[TEST] Sending request to /generate...');

const req = http.request({
  hostname: 'localhost',
  port: 3000,
  path: '/generate',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(data)
  }
}, (res) => {
  let body = '';
  res.on('data', (chunk) => { body += chunk; });
  res.on('end', () => {
    console.log('[TEST] Response status:', res.statusCode);
    try {
      const json = JSON.parse(body);
      console.log('[TEST] Response JSON:');
      console.log('  demo:', json.demo);
      console.log('  fallback:', json.fallback);
      console.log('  savedImagePath:', json.savedImagePath);
      console.log('  imageLength:', String(json.image || '').length);
      console.log('  imagePrefix:', String(json.image || '').substring(0, 50));
    } catch (e) {
      console.log('[TEST] Response body (first 500 chars):', body.substring(0, 500));
    }
    process.exit(0);
  });
});

req.on('error', (e) => {
  console.log('[TEST] Request error:', e.message);
  process.exit(1);
});

req.write(data);
req.end();
