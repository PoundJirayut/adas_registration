require('dotenv').config();
const express = require('express');
const axios   = require('axios');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const SHEET_ID   = process.env.SHEET_ID  || '19TspRNs1fkeY89CP-lJW8sdXaTCSGUeD8wh8FPpKoww';
const SCRIPT_URL = process.env.SCRIPT_URL || '';
const PORT       = process.env.PORT       || 3000;
const BASE_URL   = process.env.BASE_URL
  || (process.env.RAILWAY_PUBLIC_DOMAIN
      ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}`
      : `http://localhost:${PORT}`);

const UPLOAD_DIR = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

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
    // ── Step 1: Update Sheet via GET (always works, no redirect issue) ──
    const params = new URLSearchParams({ action: 'register', paperId });
    const regResp = await axios.get(`${SCRIPT_URL}?${params}`);
    const regData = regResp.data;

    if (!regData.success) {
      return res.json(regData);
    }

    // ── Step 2: Upload image to Drive via doPost (separate call) ────────
    let driveUrl   = '';
    let driveError = '';

    if (imageBase64) {
      const isLocalhost = BASE_URL.includes('localhost') || BASE_URL.includes('127.0.0.1');

      if (isLocalhost) {
        // Save locally — Apps Script can't reach localhost, skip Drive upload
        const filename = `${paperId}_${Date.now()}.jpg`;
        fs.writeFileSync(path.join(UPLOAD_DIR, filename), Buffer.from(imageBase64, 'base64'));
        driveError = 'Running on localhost — Drive upload skipped (deploy to Railway to enable)';
        console.warn('[/api/register] Drive upload skipped (localhost)');
      } else {
        // On Railway: send base64 directly to Apps Script doPost
        // Apps Script exec URL returns 302 on POST; we follow the redirect while keeping POST+body
        try {
          const payload = JSON.stringify({ action: 'uploadImage', paperId, imageBase64 });

          let uploadResp = await axios.post(SCRIPT_URL, payload, {
            headers: { 'Content-Type': 'application/json' },
            maxRedirects: 0,
            validateStatus: () => true,
          });

          if ((uploadResp.status === 301 || uploadResp.status === 302) && uploadResp.headers.location) {
            uploadResp = await axios.post(uploadResp.headers.location, payload, {
              headers: { 'Content-Type': 'application/json' },
              validateStatus: () => true,
            });
          }

          if (uploadResp.data?.driveUrl) {
            driveUrl = uploadResp.data.driveUrl;
            // Write Drive URL back to Sheet
            await axios.get(`${SCRIPT_URL}?${new URLSearchParams({ action: 'setPicture', paperId, driveUrl })}`);
            console.log('[/api/register] Drive upload OK:', driveUrl);
          } else {
            driveError = uploadResp.data?.error || 'Drive upload failed';
            console.warn('[/api/register] Drive upload warning:', driveError);
          }
        } catch (uploadErr) {
          driveError = uploadErr.message;
          console.warn('[/api/register] Drive upload error:', driveError);
        }
      }
    }

    res.json({ ...regData, driveUrl, driveError: driveError || undefined });
  } catch (err) {
    console.error('[/api/register]', err.message);
    res.status(500).json({ success: false, message: err.message });
  }
});

// ── Start ──────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n  APSIPA Registration running at http://localhost:${PORT}`);
  console.log(`  BASE_URL: ${BASE_URL}\n`);
  if (!SCRIPT_URL) {
    console.warn('  ⚠  SCRIPT_URL not configured — registration will not work\n');
  }
});
