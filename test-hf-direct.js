#!/usr/bin/env node
require('dotenv').config();
const https = require('https');

const hfToken = process.env.HUGGINGFACE_API_KEY;
if (!hfToken) {
  console.log('[TEST] ERROR: HUGGINGFACE_API_KEY is missing');
  process.exit(1);
}

console.log('[TEST] Token length:', hfToken.length);
console.log('[TEST] Testing Hugging Face direct API...\n');

const payload = JSON.stringify({
  inputs: 'Front view professional stage with neon LED',
  parameters: {
    negative_prompt: 'blurry, low quality'
  }
});

const models = [
  'stabilityai/stable-diffusion-xl-base-1.0',
  'black-forest-labs/FLUX.1-schnell'
];

async function testModel(model) {
  return new Promise((resolve) => {
    const options = {
      hostname: 'router.huggingface.co',
      path: `/hf-inference/models/${model}`,
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${hfToken}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload)
      }
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        console.log(`[${model}] Status: ${res.statusCode}`);
        if (res.statusCode !== 200) {
          console.log(`[${model}] Response: ${body.slice(0, 200)}`);
        } else {
          console.log(`[${model}] Success! Image size: ${body.length} bytes`);
        }
        resolve();
      });
    });

    req.on('error', (err) => {
      console.log(`[${model}] ERROR: ${err.message}`);
      resolve();
    });

    req.write(payload);
    req.end();
  });
}

(async () => {
  for (const model of models) {
    await testModel(model);
    console.log();
  }
  process.exit(0);
})();
