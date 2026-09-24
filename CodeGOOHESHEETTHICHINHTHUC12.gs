/**
 * Google Sheet -> JSON API (Apps Script Web App)
 *
 * GET  <link /exec>                 -> trả về toàn bộ dữ liệu sheet dạng JSON
 * GET  <link /exec>?sheet=TenSheet  -> lấy sheet theo tên
 * POST <link /exec>                 -> thêm dòng mới vào sheet
 *      body: {"row": {"Ten": "An", "Tuoi": 20}}
 *         hoặc {"rows": [{...}, {...}]}
 *      (có thể kèm "sheet": "TenSheet" và "token": "...")
 *
 * Dòng 1 của sheet là tiêu đề cột; key trong JSON phải trùng tiêu đề cột.
 */

const CONFIG = {
  SHEET_ID: "",      // Để trống nếu script được tạo từ Tiện ích mở rộng → Apps Script trong chính file Sheet.
                     // Nếu là script độc lập, dán ID của Sheet vào đây.
  DEFAULT_SHEET: "", // Tên sheet mặc định. Để trống = sheet đầu tiên.
  WRITE_TOKEN: ""    // Chuỗi bí mật để bảo vệ doPost. Để trống = không kiểm tra.
};

function getSpreadsheet_() {
  return CONFIG.SHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const target = name || CONFIG.DEFAULT_SHEET;
  const sheet = target ? ss.getSheetByName(target) : ss.getSheets()[0];
  if (!sheet) throw new Error("Không tìm thấy sheet: " + target);
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- ĐỌC DỮ LIỆU ----------
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const sheet = getSheet_(params.sheet);
    const values = sheet.getDataRange().getValues();

    if (values.length < 2) return json_({ ok: true, count: 0, data: [] });

    const headers = values[0].map(String);
    const data = values
      .slice(1)
      .filter(row => row.some(cell => cell !== ""))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => (obj[h] = row[i]));
        return obj;
      });

    return json_({ ok: true, count: data.length, data: data });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- GHI DỮ LIỆU ----------
function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(10000);
    locked = true;

    const body = JSON.parse(e.postData.contents);

    if (CONFIG.WRITE_TOKEN && body.token !== CONFIG.WRITE_TOKEN) {
      return json_({ ok: false, error: "Sai token" });
    }

    const sheet = getSheet_(body.sheet);
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(String);

    const records = Array.isArray(body.rows) ? body.rows : [body.row];
    const rows = records.map(rec =>
      headers.map(h => (rec && rec[h] !== undefined ? rec[h] : ""))
    );

    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
      .setValues(rows);

    return json_({ ok: true, added: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}
