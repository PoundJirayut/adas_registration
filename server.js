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
  || (process.env.RAILWAY_PUBLIC_DOMAIN ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}` : `http://localhost:${PORT}`);

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
    return res.status(500).json({ success: false, message: 'SCRIPT_URL ยังไม่ได้ตั้งค่าใน .env' });
  }

  try {
    // ── 1. บันทึกรูปไว้ที่ public/uploads/ ──────────────────────
    let imageUrl = '';
    if (imageBase64) {
      const filename = `${paperId}_${Date.now()}.jpg`;
      const filepath = path.join(UPLOAD_DIR, filename);
      fs.writeFileSync(filepath, Buffer.from(imageBase64, 'base64'));
      imageUrl = `${BASE_URL}/uploads/${filename}`;
    }

    // ── 2. เรียก Apps Script ผ่าน GET (เชื่อถือได้, ไม่มี redirect ปัญหา) ──
    const scriptParams = new URLSearchParams({ action: 'register', paperId });
    if (imageUrl) scriptParams.set('imageUrl', imageUrl);

    const resp = await axios.get(`${SCRIPT_URL}?${scriptParams.toString()}`);
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
  console.log(`\n  ADAS Registration running at http://localhost:${PORT}\n`);
  if (!SCRIPT_URL) {
    console.warn('  ⚠  SCRIPT_URL ยังไม่ได้ตั้งค่า — การลงทะเบียนจะยังไม่ทำงาน\n');
  }
});
