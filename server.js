const express    = require('express');
const multer     = require('multer');
const cors       = require('cors');
const ExcelJS    = require('exceljs');
const TelegramBot = require('node-telegram-bot-api');
const os         = require('os');
const fs         = require('fs');
const path       = require('path');
const crypto     = require('crypto');

const { fromPDF, fromImage, fromExcel, fromZIP, SUPPORTED_IMG_EXTS } = require('./extractor');

const TOKEN = '8257328752:AAFF_1_BfR4YBbzA2nT4NkLuy2p5qB-u8x4';
// Bot instance for sending files back to users (no polling — bot.js handles that)
const bot = new TelegramBot(TOKEN);

// Temporary file store: token → { filePath, expires }
const downloadStore = new Map();

const app  = express();
const PORT = process.env.PORT || 3000;

// ─── Multer — temp file upload ────────────────────────────────────────────────
const upload = multer({
  dest: os.tmpdir(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const ok  = ['.pdf', '.xlsx', '.xls', '.zip', ...SUPPORTED_IMG_EXTS];
    if (ok.includes(ext)) cb(null, true);
    else cb(new Error(`Unsupported file type: ${ext}`));
  },
});

// ─── Excel column definition ──────────────────────────────────────────────────
const EXCEL_COLUMNS = [
  { header: 'AssetNo',                  key: 'assetNo',                 width: 15 },
  { header: 'VehicleType',              key: 'vehicleType',             width: 14 },
  { header: 'CompanyName',              key: 'companyName',             width: 35 },
  { header: 'InstallationType',         key: 'installationType',        width: 18 },
  { header: 'Asateel',                  key: 'asateel',                 width: 10 },
  { header: 'Securepath',               key: 'securepath',              width: 12 },
  { header: 'InstallationCertificate',  key: 'installationCertificate', width: 24 },
  { header: 'SecurepathPlus',           key: 'securepathPlus',          width: 16 },
  { header: 'Shaheen',                  key: 'shaheen',                 width: 10 },
  { header: 'ChassisNo',                key: 'chassisNo',               width: 20 },
  { header: 'UserID',                   key: 'userId',                  width: 12 },
  { header: 'Pass',                     key: 'pass',                    width: 12 },
  { header: 'IMEI',                     key: 'imei',                    width: 18 },
];

function toRow(v) {
  return {
    assetNo:                 v.assetNo                 || '',
    vehicleType:             v.vehicleType             || '',
    companyName:             v.companyName             || '',
    installationType:        v.installationType        || 'Nil',
    asateel:                 v.asateel                 || 'Nil',
    securepath:              v.securepath              || 'Nil',
    installationCertificate: v.installationCertificate || 'Nil',
    securepathPlus:          v.securepathPlus          || 'Nil',
    shaheen:                 v.shaheen                 || 'Nil',
    chassisNo:               v.chassisNo               || '',
    userId:                  v.userId                  || '',
    pass:                    v.pass                    || '',
    imei:                    v.imei                    || '',
  };
}

async function buildExcel(vehicles) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mulkia Bot';
  wb.created = new Date();
  const ws = wb.addWorksheet('data');
  ws.columns = EXCEL_COLUMNS;

  // Header style
  const hr = ws.getRow(1);
  hr.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  hr.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F497D' } };
  hr.alignment = { vertical: 'middle', horizontal: 'center' };
  hr.height    = 20;

  for (const v of vehicles) {
    const row = ws.addRow(toRow(v));
    row.alignment = { vertical: 'middle' };
    if (row.number % 2 === 0) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    }
  }

  // Borders
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= EXCEL_COLUMNS.length; c++) {
      ws.getCell(r, c).border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
    }
  }
  ws.autoFilter = { from: 'A1', to: `M${ws.rowCount}` };

  const outPath = path.join(os.tmpdir(), `mulkia_export_${Date.now()}.xlsx`);
  await wb.xlsx.writeFile(outPath);
  return outPath;
}

function cleanupFile(f) {
  try { if (f && fs.existsSync(f)) fs.unlinkSync(f); } catch (_) {}
}

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── API: Extract from uploaded file ─────────────────────────────────────────
app.post('/api/extract', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const filePath = req.file.path;
  const ext      = path.extname(req.file.originalname).toLowerCase();
  const label    = req.file.originalname;

  try {
    let vehicles = [], stats = {}, errors = [];

    if (ext === '.pdf') {
      const { vehicle, ocrUsed } = await fromPDF(filePath, label);
      vehicles = [vehicle];
      stats = { pdf: 1, ocrUsed };

    } else if (SUPPORTED_IMG_EXTS.includes(ext)) {
      const { vehicle } = await fromImage(filePath, label);
      vehicles = [vehicle];
      stats = { image: 1 };

    } else if (ext === '.xlsx' || ext === '.xls') {
      vehicles = fromExcel(filePath);
      stats = { excel: vehicles.length };

    } else if (ext === '.zip') {
      const result = await fromZIP(filePath);
      vehicles = result.vehicles;
      stats    = result.stats;
      errors   = result.errors;
    }

    return res.json({ vehicles, stats, errors });

  } catch (e) {
    return res.status(422).json({ error: e.message });
  } finally {
    cleanupFile(filePath);
  }
});

// ─── API: Export → send file to Telegram chat ────────────────────────────────
// Used by Mini App when running inside Telegram (has a user ID)
app.post('/api/export-to-chat', async (req, res) => {
  const { vehicles, userId } = req.body;
  if (!vehicles?.length)      return res.status(400).json({ error: 'No vehicle data' });
  if (!userId)                return res.status(400).json({ error: 'No user ID' });

  let excelPath;
  try {
    excelPath = await buildExcel(vehicles);
    const filename = `mulkia_${new Date().toISOString().slice(0,10)}.xlsx`;
    await bot.sendDocument(String(userId), excelPath, {
      caption: `📊 *${vehicles.length} vehicle(s) extracted*`,
      parse_mode: 'Markdown',
    }, { filename });
    res.json({ ok: true });
  } catch (e) {
    console.error('export-to-chat error:', e.message);
    res.status(500).json({ error: e.message });
  } finally {
    cleanupFile(excelPath);
  }
});

// ─── API: Export → generate temp download token (fallback for browser) ───────
app.post('/api/export', async (req, res) => {
  const { vehicles } = req.body;
  if (!vehicles?.length) return res.status(400).json({ error: 'No vehicle data' });

  let excelPath;
  try {
    excelPath = await buildExcel(vehicles);
    // Store with a 5-minute expiry token
    const token = crypto.randomBytes(20).toString('hex');
    downloadStore.set(token, {
      filePath: excelPath,
      filename: `mulkia_${new Date().toISOString().slice(0,10)}.xlsx`,
      expires:  Date.now() + 5 * 60 * 1000,
    });
    // Cleanup token after expiry
    setTimeout(() => {
      const entry = downloadStore.get(token);
      if (entry) { cleanupFile(entry.filePath); downloadStore.delete(token); }
    }, 5 * 60 * 1000);

    res.json({ token });
  } catch (e) {
    cleanupFile(excelPath);
    res.status(500).json({ error: e.message });
  }
});

// ─── GET /api/download/:token — actual file download ─────────────────────────
app.get('/api/download/:token', (req, res) => {
  const entry = downloadStore.get(req.params.token);
  if (!entry) return res.status(404).send('Download link expired or not found.');
  if (Date.now() > entry.expires) {
    cleanupFile(entry.filePath);
    downloadStore.delete(req.params.token);
    return res.status(410).send('Download link has expired.');
  }
  res.setHeader('Content-Disposition', `attachment; filename="${entry.filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const stream = fs.createReadStream(entry.filePath);
  stream.pipe(res);
  stream.on('end', () => {
    cleanupFile(entry.filePath);
    downloadStore.delete(req.params.token);
  });
  stream.on('error', () => { cleanupFile(entry.filePath); downloadStore.delete(req.params.token); res.status(500).end(); });
});

// ─── Error handler for multer ─────────────────────────────────────────────────
app.use((err, req, res, next) => {
  if (err.code === 'LIMIT_FILE_SIZE') return res.status(413).json({ error: 'File too large (max 50 MB)' });
  if (err.message) return res.status(400).json({ error: err.message });
  next(err);
});

app.listen(PORT, () => console.log(`🌐 Mini App server running on http://localhost:${PORT}`));

module.exports = app;
