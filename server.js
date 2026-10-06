require('dotenv').config();
const path = require('path');
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');

const app = express();
app.use(express.json({ limit: '15mb' }));
app.use(express.static(path.join(__dirname, 'public')));

if (!process.env.ANTHROPIC_API_KEY) {
  console.warn('ANTHROPIC_API_KEY manquante — copie .env.example en .env et renseigne ta cle.');
}

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const MODELS = {
  quick: process.env.MODEL_QUICK || 'claude-haiku-4-5-20251001',
  default: process.env.MODEL_DEFAULT || 'claude-sonnet-5'
};

function errorCode(err) {
  const status = err && (err.status || (err.error && err.error.status));
  if (status === 429) return 'rate_limited';
  if (status === 401 || status === 403) return 'not_granted';
  if (status === 400 && String(err.message || '').toLowerCase().includes('image')) return 'image_rejected';
  return 'refused';
}

function extractText(message) {
  return (message.content || [])
    .filter((b) => b.type === 'text')
    .map((b) => b.text)
    .join('')
    .trim();
}

app.post('/api/ask', async (req, res) => {
  try {
    const { prompt, modelTier, image } = req.body || {};
    if (!prompt) return res.status(400).json({ code: 'refused' });

    const content = [];
    if (image && image.data) {
      content.push({
        type: 'image',
        source: { type: 'base64', media_type: image.mediaType || 'image/jpeg', data: image.data }
      });
    }
    content.push({ type: 'text', text: prompt });

    const message = await anthropic.messages.create({
      model: MODELS[modelTier] || MODELS.default,
      max_tokens: 1536,
      messages: [{ role: 'user', content }]
    });

    const text = extractText(message);
    if (!text) return res.status(422).json({ code: 'empty_completion' });
    res.json({ text });
  } catch (err) {
    console.error('POST /api/ask', err);
    res.status(500).json({ code: errorCode(err) });
  }
});

app.post('/api/ask-json', async (req, res) => {
  try {
    const { prompt, modelTier, image } = req.body || {};
    if (!prompt) return res.status(400).json({ code: 'refused' });

    const content = [];
    if (image && image.data) {
      content.push({
        type: 'image',
        source: { type: 'base64', media_type: image.mediaType || 'image/jpeg', data: image.data }
      });
    }
    content.push({ type: 'text', text: prompt });

    const message = await anthropic.messages.create({
      model: MODELS[modelTier] || MODELS.default,
      max_tokens: 1536,
      messages: [{ role: 'user', content }]
    });

    const text = extractText(message);
    if (!text) return res.status(422).json({ code: 'empty_completion' });

    let jsonText = text;
    const fence = jsonText.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence) jsonText = fence[1];

    let result;
    try {
      result = JSON.parse(jsonText);
    } catch (parseErr) {
      return res.status(422).json({ code: 'invalid_json' });
    }
    res.json({ result });
  } catch (err) {
    console.error('POST /api/ask-json', err);
    res.status(500).json({ code: errorCode(err) });
  }
});

app.post('/api/import-stock', (req, res) => {
  try {
    const { fileBase64 } = req.body || {};
    if (!fileBase64) return res.status(400).json({ code: 'refused' });
    const buffer = Buffer.from(fileBase64, 'base64');
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    const toIsoDate = (val) => {
      if (!val && val !== 0) return null;
      if (val instanceof Date) return val.toISOString().slice(0, 10);
      const s = String(val).trim();
      if (!s) return null;
      const d = new Date(s);
      return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
    };

    const items = rows
      .slice(1)
      .map((row) => ({
        name: String(row[0] || '').trim(),
        qty: row[1] !== undefined ? String(row[1]).trim() : '',
        unit: String(row[2] || '').trim(),
        expiry: toIsoDate(row[3])
      }))
      .filter((item) => item.name);

    res.json({ items });
  } catch (err) {
    console.error('POST /api/import-stock', err);
    res.status(500).json({ code: 'refused' });
  }
});

const PORT = process.env.PORT || 3100;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Labo tourne sur http://localhost:${PORT}`);
});
