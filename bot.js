const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');
const os = require('os');

const { fromPDF, fromImage, fromExcel, fromZIP, SUPPORTED_IMG_EXTS } = require('./extractor');

// ─── Config ───────────────────────────────────────────────────────────────────
const TOKEN    = '8257328752:AAFF_1_BfR4YBbzA2nT4NkLuy2p5qB-u8x4';
const APP_URL  = process.env.MINI_APP_URL || 'http://localhost:3000'; // set MINI_APP_URL for production
const MAX_MB   = 50;

const SUPPORTED_DOC_EXTS = ['.pdf', '.xlsx', '.xls', '.zip'];
const ALL_EXTS = [...SUPPORTED_DOC_EXTS, ...SUPPORTED_IMG_EXTS];

// ─── Excel builder (same as server.js) ───────────────────────────────────────
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
    assetNo: v.assetNo || '', vehicleType: v.vehicleType || '',
    companyName: v.companyName || '',
    installationType: v.installationType || 'Nil', asateel: v.asateel || 'Nil',
    securepath: v.securepath || 'Nil', installationCertificate: v.installationCertificate || 'Nil',
    securepathPlus: v.securepathPlus || 'Nil', shaheen: v.shaheen || 'Nil',
    chassisNo: v.chassisNo || '', userId: v.userId || '', pass: v.pass || '', imei: v.imei || '',
  };
}

async function buildExcel(vehicles) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mulkia Bot'; wb.created = new Date();
  const ws = wb.addWorksheet('data');
  ws.columns = EXCEL_COLUMNS;
  const hr = ws.getRow(1);
  hr.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  hr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F497D' } };
  hr.alignment = { vertical: 'middle', horizontal: 'center' }; hr.height = 20;
  for (const v of vehicles) {
    const row = ws.addRow(toRow(v));
    row.alignment = { vertical: 'middle' };
    if (row.number % 2 === 0) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
  }
  for (let r = 1; r <= ws.rowCount; r++)
    for (let c = 1; c <= EXCEL_COLUMNS.length; c++)
      ws.getCell(r, c).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  ws.autoFilter = { from: 'A1', to: `M${ws.rowCount}` };
  const out = path.join(os.tmpdir(), `mulkia_${Date.now()}.xlsx`);
  await wb.xlsx.writeFile(out); return out;
}

// ─── Bot init ─────────────────────────────────────────────────────────────────
let bot;
try {
  bot = new TelegramBot(TOKEN, { polling: true });
  console.log('✅ Bot started.');
} catch (e) { console.error('FATAL:', e.message); process.exit(1); }

// ─── Helpers ──────────────────────────────────────────────────────────────────
function safeMsg(chatId, text, opts = {}) {
  const t = text.length > 4000 ? text.substring(0, 4000) + '\n_[truncated]_' : text;
  return bot.sendMessage(chatId, t, opts).catch(e => console.error('sendMessage:', e.message));
}
function tmpFile(ext = '') { return path.join(os.tmpdir(), `mulkia_${Date.now()}_${Math.random().toString(36).slice(2)}${ext}`); }
function cleanup(f) { try { if (f && fs.existsSync(f)) fs.unlinkSync(f); } catch (_) {} }

async function sendExcel(chatId, vehicles, caption) {
  let p;
  try {
    p = await buildExcel(vehicles);
    await bot.sendDocument(chatId, p, { caption, parse_mode: 'Markdown' });
  } catch (e) { await safeMsg(chatId, `❌ Failed to generate Excel: ${e.message}`); }
  finally { cleanup(p); }
}

function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith('https') ? https : http;
    const file = fs.createWriteStream(dest);
    let settled = false;
    const done = (err) => {
      if (settled) return; settled = true;
      if (err) { file.destroy(); fs.unlink(dest, () => {}); reject(err); }
      else file.close(() => resolve(dest));
    };
    const req = proto.get(url, res => {
      if (res.statusCode !== 200) return done(new Error(`HTTP ${res.statusCode}`));
      res.pipe(file);
      file.on('finish', () => done(null));
      file.on('error', e => done(new Error('Write: ' + e.message)));
    });
    req.on('error', e => done(new Error('Net: ' + e.message)));
    req.setTimeout(120000, () => { req.destroy(); done(new Error('Timed out')); });
  });
}

// ─── /start ───────────────────────────────────────────────────────────────────
bot.onText(/\/start/, async (msg) => {
  const name = msg.from?.first_name || 'there';
  await bot.sendMessage(msg.chat.id,
    `👋 Hello *${name}*! I'm the *Mulkia Bot* 🚗\n\n` +
    `Choose how to use me:\n\n` +
    `  🌐 *Open the Mini App* for a full interface\n` +
    `  📤 Or just send files directly here\n\n` +
    `_Supported: PDF · Photo · ZIP · Excel_`,
    {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [[
          { text: '🌐 Open Mini App', web_app: { url: APP_URL } }
        ]]
      }
    }
  ).catch(e => console.error('start msg:', e.message));
});

// ─── File upload handler ──────────────────────────────────────────────────────
async function handleUpload(msg) {
  const chatId = msg.chat.id;
  let fileId, fileName, fileSize, isPhoto = false;

  if (msg.photo) {
    const p = msg.photo[msg.photo.length - 1];
    fileId = p.file_id; fileSize = p.file_size || 0;
    fileName = `photo_${Date.now()}.jpg`; isPhoto = true;
  } else if (msg.document) {
    fileId = msg.document.file_id;
    fileName = msg.document.file_name || `file_${Date.now()}`;
    fileSize = msg.document.file_size || 0;
  } else return;

  const ext = isPhoto ? '.jpg' : path.extname(fileName).toLowerCase();
  if (!isPhoto && !ALL_EXTS.includes(ext)) {
    return safeMsg(chatId,
      `❌ Unsupported file type \`${ext}\`\n\nSupported: PDF · JPG/PNG/BMP/TIFF/WEBP · ZIP · XLSX/XLS`,
      { parse_mode: 'Markdown' }
    );
  }

  const sizeMB = fileSize / (1024 * 1024);
  if (sizeMB > MAX_MB) return safeMsg(chatId, `❌ File too large: ${sizeMB.toFixed(1)} MB (max ${MAX_MB} MB)`);

  let fileInfo;
  try {
    fileInfo = await bot.getFile(fileId);
    if (!fileInfo?.file_path) throw new Error('Missing file_path');
  } catch (e) {
    return safeMsg(chatId, `❌ Cannot get file from Telegram: ${e.message}`);
  }

  await safeMsg(chatId,
    SUPPORTED_IMG_EXTS.includes(ext)
      ? `📥 Received photo\n🔍 Running OCR... _(10–30 sec)_`
      : `📥 Received *${fileName}*\n⚙️ Processing...`,
    { parse_mode: 'Markdown' }
  );

  const dest = tmpFile(ext);
  try {
    await downloadFile(`https://api.telegram.org/file/bot${TOKEN}/${fileInfo.file_path}`, dest);
  } catch (e) { cleanup(dest); return safeMsg(chatId, `❌ Download failed: ${e.message}`); }

  if (!fs.existsSync(dest) || fs.statSync(dest).size === 0) {
    cleanup(dest); return safeMsg(chatId, `❌ Downloaded file is empty. Please re-send.`);
  }

  try {
    if (ext === '.pdf') {
      let result;
      try { result = await fromPDF(dest, fileName); }
      catch (e) {
        return safeMsg(chatId,
          `❌ *PDF Error*\n\n${e.message}\n\n💡 Ensure PDF has selectable text or is a clear scan. Not password-protected.`
        );
      }
      await sendExcel(chatId, [result.vehicle],
        result.ocrUsed ? `✅ Mulkia extracted via OCR` : `✅ Mulkia extracted — 1 vehicle`
      );

    } else if (SUPPORTED_IMG_EXTS.includes(ext)) {
      let result;
      try { result = await fromImage(dest, fileName); }
      catch (e) {
        return safeMsg(chatId,
          `❌ *OCR Error*\n\n${e.message}\n\n💡 Use a clear, well-lit photo with the full mulkia visible.`
        );
      }
      await sendExcel(chatId, [result.vehicle], `✅ Mulkia extracted from image`);

    } else if (ext === '.xlsx' || ext === '.xls') {
      let vehicles;
      try { vehicles = fromExcel(dest); }
      catch (e) {
        return safeMsg(chatId,
          `❌ *Excel Error*\n\n${e.message}\n\n💡 File needs "AssetNo" or "ChassisNo" column.`
        );
      }
      await sendExcel(chatId, vehicles, `✅ Excel re-exported — ${vehicles.length} vehicle(s)`);

    } else if (ext === '.zip') {
      await safeMsg(chatId, `⚙️ Processing ZIP...`);
      let result;
      try { result = await fromZIP(dest); }
      catch (e) {
        return safeMsg(chatId, `❌ *ZIP Error*\n\n${e.message}\n\n💡 ZIP must not be password-protected.`);
      }

      const { vehicles, stats, errors } = result;
      if (!vehicles.length) {
        let msg2 = `⚠️ ZIP processed but no vehicles extracted.`;
        if (errors.length) msg2 += `\n\n*Errors:*\n` + errors.slice(0,5).map(e=>`• ${e.file}: ${e.error}`).join('\n');
        return safeMsg(chatId, msg2, { parse_mode: 'Markdown' });
      }

      let summary = `✅ *ZIP processed!*\n📄 ${stats.pdf} PDF · 🖼️ ${stats.image} Images · 📊 ${stats.excel} Excel rows`;
      if (stats.failed) summary += ` · ❌ ${stats.failed} failed`;
      if (errors.length) {
        summary += `\n\n⚠️ *Errors:*\n` + errors.slice(0,3).map(e=>`• ${e.file}: ${e.error}`).join('\n');
        if (errors.length > 3) summary += `\n_...${errors.length-3} more_`;
      }
      await safeMsg(chatId, summary, { parse_mode: 'Markdown' });
      await sendExcel(chatId, vehicles, `📊 *${vehicles.length} vehicle(s) extracted*`);
    }

  } finally { cleanup(dest); }
}

bot.on('document', handleUpload);
bot.on('photo',    handleUpload);

bot.on('message', async (msg) => {
  if (msg.document || msg.photo || !msg.text) return;
  if (msg.text.startsWith('/')) return;
  await safeMsg(msg.chat.id,
    `📤 Send a PDF, photo, ZIP, or Excel file — I'll extract the vehicle data and return an Excel.\n\nOr use /start to open the Mini App.`
  );
});

bot.on('polling_error', err => {
  console.error('Polling error:', err.code, err.message);
  if (err.code === 'EFATAL') setTimeout(() => process.exit(1), 10000);
});
bot.on('error', err => console.error('Bot error:', err.message));
process.on('uncaughtException',  err => console.error('Uncaught:', err.message, err.stack));
process.on('unhandledRejection', r   => console.error('Unhandled rejection:', r));
process.on('SIGINT', () => { bot.stopPolling(); process.exit(0); });

console.log('🤖 Mulkia Bot running.');
console.log(`🌐 Mini App URL: ${APP_URL}`);
