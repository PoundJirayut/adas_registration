require('dotenv').config();
const express = require('express');
const axios   = require('axios');
const path    = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const SHEET_ID   = process.env.SHEET_ID  || '19TspRNs1fkeY89CP-lJW8sdXaTCSGUeD8wh8FPpKoww';
const SCRIPT_URL = process.env.SCRIPT_URL || '';
const PORT       = process.env.PORT       || 3000;

// ── GET /api/data ──────────────────────────────────────────────
app.get('/api/data', async (req, res) => {
  try {
    const url  = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json`;
    const resp = await axios.get(url, { responseType: 'text' });
    const text = resp.data;

    const json = JSON.parse(text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1));
    const { cols, rows } = json.table;

    let headers  = cols.map(c => c.label);
    let startRow = 0;
    if (headers.every(h => !h) && rows?.length) {
      headers  = (rows[0].c || []).map(cell => cell?.v ?? '');
      startRow = 1;
    }

    const data = (rows || []).slice(startRow).map(row =>
      Object.fromEntries((row.c || []).map((cell, i) => [headers[i], cell?.v ?? '']))
    );

    res.json({ success: true, data });
  } catch (err) {
    console.error('[/api/data]', err.message);
    res.status(500).json({ success: false, message: err.message });
  }
});

// ── POST /api/register ─────────────────────────────────────────
app.post('/api/register', async (req, res) => {
  const { paperId, imageBase64 } = req.body;

  if (!paperId) {
    return res.status(400).json({ success: false, message: 'paperId is required' });
  }
  if (!SCRIPT_URL) {
    return res.status(500).json({ success: false, message: 'SCRIPT_URL not configured in .env' });
  }

  try {
    // POST base64 directly to Apps Script — no local file, no public URL needed
    const payload = JSON.stringify({ action: 'register', paperId, imageBase64: imageBase64 || '' });

    // Apps Script exec URL returns 302 on POST; follow redirect while keeping POST + body
    let resp = await axios.post(SCRIPT_URL, payload, {
      headers: { 'Content-Type': 'application/json' },
      maxRedirects: 0,
      validateStatus: () => true,
    });

    if (resp.status === 301 || resp.status === 302) {
      const redirectUrl = resp.headers.location;
      resp = await axios.post(redirectUrl, payload, {
        headers: { 'Content-Type': 'application/json' },
      });
    }

    const data = resp.data;
    if (data.driveError) {
      console.warn('[/api/register] Drive upload warning:', data.driveError);
    } else if (data.driveUrl) {
      console.log('[/api/register] Drive upload OK:', data.driveUrl);
    }
    res.json(data);
  } catch (err) {
    console.error('[/api/register]', err.message);
    res.status(500).json({ success: false, message: err.message });
  }
});

// ── Start ──────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n  APSIPA Registration running at http://localhost:${PORT}\n`);
  if (!SCRIPT_URL) {
    console.warn('  ⚠  SCRIPT_URL not configured — registration will not work\n');
  }
});
