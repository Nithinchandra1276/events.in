const https = require('https');

function fetchUrl(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return resolve(fetchUrl(res.headers.location));
      }

      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        const buffer = Buffer.concat(chunks);
        resolve({
          status: res.statusCode,
          contentType: res.headers['content-type'] || '',
          size: buffer.length,
          isJpeg: buffer[0] === 0xff && buffer[1] === 0xd8,
          isPng: buffer[0] === 0x89 && buffer[1] === 0x50,
        });
      });
    }).on('error', reject);
  });
}

(async () => {
  const url = 'https://image.pollinations.ai/prompt/front%20view%20event%20stage%20with%20led%20wall%20and%20speakers%20high%20quality?width=1024&height=768&nologo=true';
  const r = await fetchUrl(url);
  console.log(r);
})();
