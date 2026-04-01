// Shared extraction logic — used by both bot.js and server.js
const pdfParse = require('pdf-parse');
const { createWorker } = require('tesseract.js');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const SUPPORTED_IMG_EXTS = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'];

// ─── PDF text extraction ──────────────────────────────────────────────────────
async function extractPDFText(filePath) {
  try {
    const text = execFileSync('pdftotext', [filePath, '-'], {
      encoding: 'utf8', timeout: 30000, maxBuffer: 10 * 1024 * 1024
    });
    if (text && text.trim().length > 10) return text;
  } catch (e) {
    console.warn('pdftotext failed, trying pdf-parse:', e.message);
  }
  const buf = fs.readFileSync(filePath);
  if (!buf.length) throw new Error('PDF file is empty');
  const data = await pdfParse(buf);
  if (!data.text || data.text.trim().length < 10) throw new Error('NO_TEXT');
  return data.text;
}

// ─── OCR ─────────────────────────────────────────────────────────────────────
async function ocrImage(imagePath) {
  let worker;
  try {
    worker = await createWorker(['eng', 'ara'], 1, { logger: () => {} });
    const { data: { text } } = await worker.recognize(imagePath);
    if (!text || text.trim().length < 5) throw new Error('OCR returned empty result');
    return text;
  } finally {
    if (worker) await worker.terminate().catch(() => {});
  }
}

// ─── Mulkia text → vehicle record ─────────────────────────────────────────────
function parseMulkiaText(rawText) {
  if (!rawText || typeof rawText !== 'string') throw new Error('No text to parse');

  const clean = rawText
    .normalize('NFKC')                                            // normalize Arabic Presentation Forms → base chars
    .replace(/[\u200f\u200e\u202a-\u202e\u2066-\u2069]/g, '')    // strip RTL/LTR Unicode marks
    .replace(/\r/g, '')
    .replace(/[ \t]+/g, ' ')
    .trim();

  if (clean.length < 10) throw new Error('Text too short');

  const joined = clean.split('\n').map(l => l.trim()).filter(Boolean).join(' ');
  const v = {};

  // Plate: P/53726  A/1234  12345
  const pm = joined.match(/\b([A-Z]{1,2}\/\d{3,6})\b/) || joined.match(/\b(\d{4,6})\b/);
  if (pm) v.assetNo = pm[1];

  // Chassis/VIN — 17 chars, no I/O/Q
  const cm = joined.match(/\b([A-HJ-NPR-Z0-9]{17})\b/i);
  if (cm) v.chassisNo = cm[1].toUpperCase();

  // Company / Owner
  const companyRe = /\b([A-Z][A-Z\s&'.,()-]{6,60}?(?:LLC|L\.L\.C\.?|EST|CO\.|LTD|CORP|INC|FZCO|PJSC|TRADING|TRANSPORT|SERVICES|GENERAL|CONTRACTING|DELIVERY|RENTAL|ESTABLISHMENT)[.]*)/gi;
  const companyHits = [...joined.matchAll(companyRe)];
  if (companyHits.length) {
    v.companyName = companyHits
      .map(m => m[1].trim().replace(/\s+/g, ' '))
      .sort((a, b) => b.length - a.length)[0];
  }

  // Vehicle type — Arabic (after NFKC normalization) + English
  const arTypeMap = {
    'فان': 'Van', 'سيارة': 'Car', 'سياره': 'Car',
    'حافلة': 'Bus', 'حافله': 'Bus', 'باص': 'Bus',
    'شاحنة': 'Truck', 'شاحنه': 'Truck',
    'بيك اب': 'Pickup', 'بيك أب': 'Pickup',
    'دراجة': 'Motorcycle', 'دراجه': 'Motorcycle',
    'مقطورة': 'Trailer', 'مقطوره': 'Trailer',
    'خصوصي': 'Car', 'نقل': 'Truck',
  };
  for (const [ar, en] of Object.entries(arTypeMap)) {
    if (joined.includes(ar)) { v.vehicleType = en; break; }
  }
  if (!v.vehicleType) {
    const m = joined.match(/\b(Van|Bus|Truck|Pickup|Motorcycle|Trailer|Tanker|Saloon|Hatchback|SUV|Coupe|Wagon)\b/i);
    if (m) v.vehicleType = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
  }
  if (!v.vehicleType) {
    if (/\bprivate\b/i.test(joined) || joined.includes('خصوصي')) v.vehicleType = 'Car';
    else if (/\bpublic\b/i.test(joined)) v.vehicleType = 'Bus';
  }

  // Make/Model
  const makes = ['TOYOTA','NISSAN','MITSUBISHI','ISUZU','FORD','HYUNDAI','KIA','HONDA',
    'MERCEDES-BENZ','MERCEDES','BMW','VOLVO','MAN','SCANIA','HINO','FUSO',
    'RENAULT','PEUGEOT','CHEVROLET','GMC','LAND ROVER','JEEP','LEXUS','INFINITI',
    'DAIHATSU','SUZUKI','MAZDA','SUBARU','VOLKSWAGEN','AUDI'];
  for (const make of makes) {
    const m = joined.match(new RegExp(`\\b(${make}[\\w\\s\\/-]{0,25})\\b`, 'i'));
    if (m) { v.makeModel = m[0].trim().replace(/\s+/g, ' '); break; }
  }

  // Dates DD-MM-YYYY or DD/MM/YYYY
  const dates = [...joined.matchAll(/\b(\d{2}[-\/]\d{2}[-\/]\d{4})\b/g)].map(m => m[1]);
  if (dates[0]) v.regDate    = dates[0];
  if (dates[1]) v.expiryDate = dates[1];
  if (dates[2]) v.insExpiry  = dates[2];

  // Year
  const ym = joined.match(/\b(19[9]\d|20[0-3]\d)\b/);
  if (ym) v.year = ym[1];

  // Origin
  const om = joined.match(/\b(Japan|Germany|USA|United States|Korea|South Korea|China|India|Sweden|France|Italy|UK)\b/i);
  if (om) v.origin = om[1];

  // Emirate
  const em = joined.match(/\b(Dubai|Abu Dhabi|Sharjah|Ajman|Fujairah|Ras Al Khaimah|Umm Al Quwain)\b/i);
  if (em) v.emirate = em[1];

  // Engine number
  const engineRe = [
    /\b(2TR[A-Z0-9]{5,10}|1GR[A-Z0-9]{5,10}|K24[A-Z0-9]{5,10}|4JJ[A-Z0-9]{5,10}|4HK[A-Z0-9]{5,10}|QR25[A-Z0-9]{4,9}|4D56[A-Z0-9]{5,10}|OM457[A-Z0-9]{3,8}|1VD[A-Z0-9]{5,10}|2KD[A-Z0-9]{5,10})\b/i,
    /(?:engine|motor)[:\s#]*([A-Z0-9]{6,15})/i,
  ];
  for (const re of engineRe) {
    const m = joined.match(re);
    if (m) { v.engineNo = (m[1] || m[0]).toUpperCase(); break; }
  }

  return v;
}

// ─── Excel → vehicles ─────────────────────────────────────────────────────────
function fromExcel(filePath) {
  let wb;
  try { wb = XLSX.readFile(filePath); }
  catch (e) { throw new Error(`Cannot open Excel: ${e.message}`); }
  if (!wb.SheetNames?.length) throw new Error('Excel file has no sheets');

  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  if (!rows?.length) throw new Error('Excel sheet has no data rows');

  const vehicles = [];
  for (const row of rows) {
    try {
      const v = { source: 'excel' };
      for (const [k, val] of Object.entries(row)) {
        const key = String(k).toLowerCase().replace(/[\s_\-()\/]/g, '');
        const sv  = val != null ? String(val).trim() : '';
        if (!sv || sv === 'undefined') continue;
        if      (key === 'assetno' || key === 'plateno' || key === 'plate') v.assetNo = sv;
        else if (key === 'chassis' || key === 'chassisno')             v.chassisNo = sv;
        else if (key.includes('companyname') || key === 'company' || key === 'owner') v.companyName = sv;
        else if (key === 'vehicletype' || key === 'type')              v.vehicleType = sv;
        else if (key === 'imei')                                        v.imei = sv;
        else if (key === 'userid')                                      v.userId = sv;
        else if (key === 'pass' || key === 'password')                  v.pass = sv;
        else if (key === 'asateel')                                     v.asateel = sv;
        else if (key === 'securepath')                                  v.securepath = sv;
        else if (key === 'installationcertificate')                     v.installationCertificate = sv;
        else if (key === 'securepathplus')                              v.securepathPlus = sv;
        else if (key === 'shaheen')                                     v.shaheen = sv;
        else if (key === 'installationtype')                            v.installationType = sv;
      }
      if (v.assetNo || v.chassisNo) vehicles.push(v);
    } catch (e) { console.error('Row error:', e.message); }
  }
  if (!vehicles.length) throw new Error('No valid records. Need "AssetNo"/"PlateNo" or "ChassisNo" column.');
  return vehicles;
}

// ─── PDF → vehicle ─────────────────────────────────────────────────────────────
async function fromPDF(filePath, label) {
  let text, ocrUsed = false;
  try {
    text = await extractPDFText(filePath);
  } catch (e) {
    if (e.message === 'NO_TEXT') { ocrUsed = true; text = await ocrImage(filePath); }
    else throw new Error(`PDF read error: ${e.message}`);
  }
  if (!text || text.trim().length < 10) throw new Error('No extractable text. PDF may be a low-resolution scan.');
  const v = parseMulkiaText(text);
  if (label) v.fileName = label;
  v.source = ocrUsed ? 'pdf-ocr' : 'pdf';
  return { vehicle: v, ocrUsed };
}

// ─── Image → vehicle ───────────────────────────────────────────────────────────
async function fromImage(filePath, label) {
  const text = await ocrImage(filePath);
  const v = parseMulkiaText(text);
  if (label) v.fileName = label;
  v.source = 'image-ocr';
  return { vehicle: v };
}

// ─── ZIP → vehicles ────────────────────────────────────────────────────────────
async function fromZIP(filePath, onProgress) {
  let zip;
  try { zip = new AdmZip(filePath); }
  catch (e) { throw new Error(`Cannot open ZIP: ${e.message}`); }

  const allExts = ['.pdf', '.xlsx', '.xls', ...SUPPORTED_IMG_EXTS];
  const entries = zip.getEntries().filter(e =>
    !e.isDirectory && allExts.includes(path.extname(e.entryName).toLowerCase())
  );
  if (!entries.length) throw new Error('ZIP has no supported files (PDF/image/Excel).');

  const tmpDir = path.join(os.tmpdir(), `mulkia_zip_${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });

  const vehicles = [];
  const errors   = [];
  const stats    = { pdf: 0, image: 0, excel: 0, failed: 0, total: entries.length };

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const ext   = path.extname(entry.entryName).toLowerCase();
    const safe  = path.basename(entry.entryName).replace(/[^a-zA-Z0-9._-]/g, '_');
    const dest  = path.join(tmpDir, safe);

    if (onProgress) onProgress(i + 1, entries.length, safe);

    try {
      zip.extractEntryTo(entry, tmpDir, false, true);
      if (!fs.existsSync(dest)) throw new Error('Missing after extraction');

      if (ext === '.pdf') {
        const { vehicle } = await fromPDF(dest, safe);
        vehicles.push(vehicle); stats.pdf++;
      } else if (SUPPORTED_IMG_EXTS.includes(ext)) {
        const { vehicle } = await fromImage(dest, safe);
        vehicles.push(vehicle); stats.image++;
      } else {
        const rows = fromExcel(dest);
        vehicles.push(...rows); stats.excel += rows.length;
      }
    } catch (e) {
      stats.failed++;
      errors.push({ file: safe, error: e.message });
    } finally {
      try { if (fs.existsSync(dest)) fs.unlinkSync(dest); } catch (_) {}
    }
  }

  try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (_) {}
  return { vehicles, stats, errors };
}

module.exports = { fromPDF, fromImage, fromExcel, fromZIP, parseMulkiaText, SUPPORTED_IMG_EXTS };
