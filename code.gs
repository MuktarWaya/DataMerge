// ============================================================
//  DataMerge — Google Apps Script Backend
//  รวมข้อมูล HDC + Mypcu (เชื่อมด้วย PID)
//
//  ผู้พัฒนา : นายมุคตาร วายา
//             พยาบาลวิชาชีพชำนาญการ  รพ.สต.ตรัง
//             Muktar32@gmail.com
// ============================================================

const SHEET_ID      = '1SDjT38DZTC5_fl9b_Zpn9P5EoKR6wnp09gN9LlcuN_o';
const LOG_SHEET     = 'LOG';
const LOG_HEADERS   = ['เวลา','ประเภท','ไฟล์ HDC','แถว HDC','ไฟล์ Mypcu','แถว Mypcu','จับคู่ได้','ไม่พบ','ไฟล์ Export','Auto-Adjust'];

// ── Serve web app ──────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DataMerge — รวมข้อมูล HDC & Mypcu')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Write log row to Google Sheet ─────────────────────────
function logUsage(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh   = ss.getSheetByName(LOG_SHEET);

    if (!sh) {
      sh = ss.insertSheet(LOG_SHEET);
      const hdr = sh.getRange(1, 1, 1, LOG_HEADERS.length);
      hdr.setValues([LOG_HEADERS]);
      hdr.setFontWeight('bold').setBackground('#f3f4f6').setFontColor('#374151');
      sh.setFrozenRows(1);
    }

    sh.appendRow([
      new Date(data.ts || Date.now()),
      data.action    || 'รวมข้อมูล',
      data.hdcFile   || '-',
      data.hdcRows   || '-',
      data.mpFile    || '-',
      data.mpRows    || '-',
      data.matched   !== undefined ? data.matched   : '-',
      data.unmatched !== undefined ? data.unmatched : '-',
      data.exportFile || '-',
      data.autoAdjust || '-',
    ]);

    return { ok: true };
  } catch (e) {
    console.error('logUsage error:', e);
    return { ok: false, error: e.toString() };
  }
}

// ── Read logs from Google Sheet (last 200 rows, newest first) ──
function getLogs() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(LOG_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];

    const lastRow  = sh.getLastRow();
    const numRows  = Math.min(lastRow - 1, 200);
    const startRow = Math.max(2, lastRow - numRows + 1);
    const values   = sh.getRange(startRow, 1, numRows, LOG_HEADERS.length).getValues();

    return values.reverse().map(r => ({
      ts:         r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      action:     r[1], hdcFile:  r[2], hdcRows:    r[3],
      mpFile:     r[4], mpRows:   r[5], matched:    r[6],
      unmatched:  r[7], exportFile: r[8], autoAdjust: r[9],
    }));
  } catch (e) {
    console.error('getLogs error:', e);
    return [];
  }
}
