require('dotenv').config();

const express = require('express');
const path = require('path');
const https = require('https');
const fs = require('fs/promises');

const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());

// Allow local frontends (3000/3001/3002 and Live Preview) to call this API.
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (!origin || /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin || '*');
  }
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.sendStatus(204);
  }
  next();
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/health', (req, res) => {
  res.json({
    ok: true,
    uptimeSeconds: Math.round(process.uptime()),
    timestamp: new Date().toISOString()
  });
});

app.get('/test-hf', async (req, res) => {
  const hfToken = process.env.HUGGINGFACE_API_KEY;  
  if (!hfToken) {
    return res.status(200).json({
      success: false,
      configured: false,
      message: 'HUGGINGFACE_API_KEY is missing'
    });
  }

  const payload = JSON.stringify({ inputs: 'test stage' });
  
  const options = {
    hostname: 'router.huggingface.co',
    path: '/hf-inference/models/stabilityai/stable-diffusion-xl-base-1.0',
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${hfToken}`,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(payload)
    }
  };

  const req2 = https.request(options, (res2) => {
    let body = '';
    res2.on('data', (c) => { body += c; });
    res2.on('end', () => {
      const parsedStatus = Number(res2.statusCode) || 0;
      res.json({
        hfToken_length: hfToken.length,
        status: parsedStatus,
        imageSize: body.length,
        success: parsedStatus === 200
      });
    });
  });
  
  req2.on('error', (err) => {
    res.json({ error: err.message });
  });
  
  req2.write(payload);
  req2.end();
});

app.use(express.static(path.join(__dirname)));
app.use('/generated-images', express.static(path.join(__dirname, 'generated-images')));

app.get('/download/ppt', (req, res) => {
  const pptPath = path.join(__dirname, 'Project_Review_22_Slides.pptx');
  return res.download(pptPath, 'Project_Review_22_Slides.pptx', (error) => {
    if (!error) return;
    console.error('Download error (/download/ppt):', error.message);
    if (!res.headersSent) {
      res.status(404).json({ error: 'PPT file not found' });
    }
  });
});

app.get('/download/ppt-minimal', (req, res) => {
  const pptPath = path.join(__dirname, 'Project_Review_22_Slides_Minimal.pptx');
  return res.download(pptPath, 'Project_Review_22_Slides_Minimal.pptx', (error) => {
    if (!error) return;
    console.error('Download error (/download/ppt-minimal):', error.message);
    if (!res.headersSent) {
      res.status(404).json({ error: 'Minimal PPT file not found' });
    }
  });
});

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;

app.post('/generate', async (req, res) => {
  try {
    process.stderr.write('[API] /generate request received\n');
    const { prompt, width, height, ledPattern, soundSystem, equipment = 'both' } = req.body;
    
    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const parsedWidth = Math.max(5, Number(width) || 20);
    const parsedHeight = Math.max(3, Number(height) || 10);

    const ledDetails = {
      geometric: 'vibrant geometric patterns in blues, purples, and whites',
      pixel: 'pixel-art style grid patterns with color transitions',
      wave: 'flowing wave patterns in cyan, magenta, and yellow',
      stage3d: 'cinematic 3D stage depth design with layered LED panels, perspective lighting, and realistic dimensional framing',
      neon: 'bright neon-colored glowing LED panels',
      hologram: '3D holographic rainbow LED effects with depth',
      spectrum: 'full spectrum rainbow colors smoothly transitioning'
    };

    const speakerCounts = {
      stereo: 2,
      quad: 4,
      surround: 6,
      line_array: 8
    };
    const speakerCount = speakerCounts[soundSystem] || speakerCounts.stereo;

    const stageFrameRule = `A real event stage in FRONT VIEW (straight camera). Show complete stage structure, stage floor, and truss frame. Stage size must look like ${parsedWidth}ft wide and ${parsedHeight}ft high.`;

    let basePrompt = "";
    if (equipment === 'both') {
      basePrompt = `${stageFrameRule} Include exactly 1 central LED wall and exactly ${speakerCount} speakers placed symmetrically on stage left and right. Include a DJ console at center front. LED style: ${ledDetails[ledPattern] || ledDetails.geometric}.`;
    } else if (equipment === 'led_only') {
      basePrompt = `${stageFrameRule} Include exactly 1 central LED wall. Do NOT include speakers. Do NOT include DJ console. LED style: ${ledDetails[ledPattern] || ledDetails.geometric}.`;
    } else if (equipment === 'sound_only') {
      basePrompt = `${stageFrameRule} Include exactly ${speakerCount} speakers and 1 DJ console. Do NOT include any LED wall or video screen.`;
    } else {
      basePrompt = `${stageFrameRule} ${prompt}`;
    }

    const enhancedPrompt = `Generate a realistic professional event stage image. ${basePrompt} Make it high clarity, sharp focus, clean lighting, crisp LED panel details, and clear speaker/DJ equipment edges. Keep object counts exact. No people, no dancers, no singers, no crowd, no text, no logo, no watermark. Do NOT generate room interior, tunnel, corridor, abstract geometry, or close-up texture.`;
    
    const { imageUrl, providerErrors } = await generateImageWithFallback(enhancedPrompt);
    let savedImagePath = null;

    if (typeof imageUrl === 'string' && imageUrl.startsWith('data:image/')) {
      try {
        savedImagePath = await saveGeneratedImage(imageUrl);
      } catch (saveError) {
        console.error('Image autosave failed:', saveError.message);
      }
    }
    
    res.json({
      image: imageUrl,
      demo: imageUrl.includes('large-outdoor-concert-stage-truss'),
      prompt: enhancedPrompt,
      savedImagePath,
      providerErrors
    });
  } catch (error) {
    console.error('Generation error:', error);
    res.status(500).json({ 
      error: error.message || 'Image generation failed',
      image: `/large-outdoor-concert-stage-truss-600nw-2695141455.webp`,
      fallback: true
    });
  }
});

async function generateImageWithFallback(prompt) {
  const providerErrors = [];

  // Try providers in this order: HuggingFace -> Gemini -> OpenAI.
  try {
    return { imageUrl: await generateWithHuggingFace(prompt), providerErrors };
  } catch (hfError) {
    providerErrors.push(`HuggingFace: ${hfError.message}`);
    process.stderr.write(`[HFHF_ERR] ${hfError.message}\n`);
    try {
      return { imageUrl: await generateWithGemini(prompt), providerErrors };
    } catch (geminiError) {
      providerErrors.push(`Gemini: ${geminiError.message}`);
      process.stderr.write(`[GEMINI_ERR] ${geminiError.message}\n`);
      try {
        return { imageUrl: await generateWithOpenAI(prompt), providerErrors };
      } catch (openAIError) {
        providerErrors.push(`OpenAI: ${openAIError.message}`);
        process.stderr.write(`[OPENAI_ERR] ${openAIError.message}\n`);
        return {
          imageUrl: `/large-outdoor-concert-stage-truss-600nw-2695141455.webp`,
          providerErrors
        };
      }
    }
  }
}

async function generateWithOpenAI(prompt) {
  if (!OPENAI_API_KEY) {
    throw new Error('OPENAI_API_KEY is missing');
  }

  return new Promise((resolve, reject) => {
    const data = JSON.stringify({
      model: 'gpt-image-1',
      prompt,
      size: '1024x1024',
      quality: 'high'
    });

    const options = {
      hostname: 'api.openai.com',
      path: '/v1/images/generations',
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json',
        'Content-Length': data.length
      }
    };

    const req = https.request(options, (res) => {
      let responseData = '';
      res.on('data', (chunk) => { responseData += chunk; });
      res.on('end', () => {
        try {
          if (res.statusCode !== 200) {
            return reject(new Error(`OpenAI API error: ${res.statusCode} - ${responseData}`));
          }

          const result = JSON.parse(responseData);
          if (result.data && result.data[0] && result.data[0].b64_json) {
            resolve(`data:image/png;base64,${result.data[0].b64_json}`);
          } else {
            reject(new Error('No image data in OpenAI response'));
          }
        } catch (e) {
          reject(e);
        }
      });
    });

    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function generateWithHuggingFace(prompt) {
  const hfToken = process.env.HUGGINGFACE_API_KEY;
  if (!hfToken) {
    throw new Error('HUGGINGFACE_API_KEY is missing');
  }

  const payload = {
    inputs: prompt,
    parameters: {
      negative_prompt: "blurry, low quality, distorted, cartoon, illustration, person, people, crowd, face, text, logo, watermark, signature, tunnel, corridor, room interior, abstract geometry, close-up texture",
      guidance_scale: 8,
      num_inference_steps: 40
    }
  };

  // Use router endpoint because legacy api-inference model URLs now return 410 Gone.
  const models = [
    'stabilityai/stable-diffusion-xl-base-1.0',
    'black-forest-labs/FLUX.1-schnell'
  ];

  process.stderr.write(`[HF_START] token len=${hfToken.length}, models=${models.length}\n`);
  let lastError = null;
  for (const model of models) {
    try {
      process.stderr.write(`[HF_TRY] ${model}\n`);
      return await requestHuggingFaceImage(model, payload, hfToken);
    } catch (err) {
      lastError = err;
      process.stderr.write(`[HF_ERR_MODEL] ${model}: ${err.message}\n`);
    }
  }

  throw lastError || new Error('All HuggingFace models failed');
}

async function generateWithGemini(prompt) {
  if (!GEMINI_API_KEY) {
    throw new Error('GEMINI_API_KEY is missing');
  }

  return new Promise((resolve, reject) => {
    const requestBody = JSON.stringify({
      contents: [
        {
          role: 'user',
          parts: [
            {
              text: prompt
            }
          ]
        }
      ],
      generationConfig: {
        responseModalities: ['IMAGE', 'TEXT']
      }
    });

    const options = {
      hostname: 'generativelanguage.googleapis.com',
      path: `/v1beta/models/gemini-2.5-flash-image:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(requestBody)
      }
    };

    const req = https.request(options, (res) => {
      let responseData = '';
      res.on('data', (chunk) => { responseData += chunk; });
      res.on('end', () => {
        try {
          if (res.statusCode !== 200) {
            return reject(new Error(`Gemini API error: ${res.statusCode} - ${responseData.slice(0, 240)}`));
          }

          const result = JSON.parse(responseData);
          const candidates = Array.isArray(result.candidates) ? result.candidates : [];

          for (const candidate of candidates) {
            const parts = candidate && candidate.content && Array.isArray(candidate.content.parts)
              ? candidate.content.parts
              : [];

            for (const part of parts) {
              if (part && part.inlineData && part.inlineData.data) {
                const mimeType = part.inlineData.mimeType || 'image/png';
                return resolve(`data:${mimeType};base64,${part.inlineData.data}`);
              }
            }
          }

          reject(new Error('No image data returned from Gemini'));
        } catch (err) {
          reject(err);
        }
      });
    });

    req.on('error', reject);
    req.write(requestBody);
    req.end();
  });
}

async function requestHuggingFaceImage(model, payload, hfToken) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload);

    const options = {
      hostname: 'router.huggingface.co',
      path: `/hf-inference/models/${model}`,
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${hfToken}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data)
      }
    };

    const req = https.request(options, (res) => {
      const chunks = [];

      res.on('data', (chunk) => {
        chunks.push(chunk);
      });

      res.on('end', () => {
        try {
          const buffer = Buffer.concat(chunks);
          if (res.statusCode !== 200) {
            const body = buffer.toString('utf8');
            return reject(new Error(`HuggingFace API error (${model}): ${res.statusCode} ${body.slice(0, 200)}`));
          }

          const contentType = String(res.headers['content-type'] || '').toLowerCase();
          const mimeType = contentType.includes('png') ? 'image/png' : 'image/jpeg';
          const base64 = buffer.toString('base64');
          resolve(`data:${mimeType};base64,${base64}`);
        } catch (e) {
          reject(e);
        }
      });
    });

    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function generateWithPicsum(prompt) {
  const seed = Math.abs(prompt.split('').reduce((a, c) => ((a << 5) - a) + c.charCodeAt(0), 0));
  return `https://picsum.photos/seed/${seed}/900/600?random=${Date.now()}`;
}

async function saveGeneratedImage(dataUrl) {
  const match = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/.exec(dataUrl);
  if (!match) {
    throw new Error('Invalid data URL format for image autosave');
  }

  const mimeType = match[1].toLowerCase();
  const base64Data = match[2];

  const extensionMap = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/webp': 'webp'
  };
  const ext = extensionMap[mimeType] || 'png';

  const outputDir = path.join(__dirname, 'generated-images');
  await fs.mkdir(outputDir, { recursive: true });

  const filename = `stage-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;
  const outputPath = path.join(outputDir, filename);
  await fs.writeFile(outputPath, Buffer.from(base64Data, 'base64'));

  return `/generated-images/${filename}`;
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/`);
});