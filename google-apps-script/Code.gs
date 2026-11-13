/**
 * Lưu kết quả các máy vào Google Sheet cho trang "Nhiệm vụ 2 – Biến hóa hình học".
 *
 * Cài đặt (một lần):
 *  1. Tạo một Google Sheet mới → Tiện ích mở rộng (Extensions) → Apps Script.
 *  2. Xoá code mẫu, dán toàn bộ file này vào, bấm Lưu.
 *  3. Triển khai (Deploy) → Tùy chọn triển khai mới (New deployment) → loại "Ứng dụng web" (Web app):
 *       - Thực thi với tư cách (Execute as): Tôi (Me)
 *       - Người có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 *     Bấm Triển khai, cấp quyền, rồi sao chép "URL ứng dụng web" (kết thúc bằng /exec).
 *  4. Dán URL đó vào SHEET_API trong index.html.
 * Khi sửa file này, phải Deploy → Quản lý triển khai → sửa (bút chì) → Phiên bản mới, để URL cũ dùng code mới.
 *
 * Mỗi dòng của trang tính "Kết quả" là một máy làm một phòng; bấm kiểm tra lại thì dòng đó được cập nhật.
 * Cột "Tên nhóm" (trường student) không còn được trang web nhập nữa: các dòng mới để trống, các dòng cũ giữ nguyên dữ liệu.
 * Cột "Thời gian làm bài" (trường duration, tính bằng giây, hiển thị dạng 1p30s) nằm sau cột "Số lần kiểm tra".
 * Trang tính cũ chưa có cột này sẽ được tự chèn thêm cột ở lần gọi tiếp theo; các dòng cũ để trống ô thời gian.
 * Cột cuối là dữ liệu trang web đọc lại (để tải PDF) — xem thoải mái nhưng đừng sửa.
 * Đừng xoá hoặc chèn cột thủ công, vì script xác định cột theo vị trí.
 */
const SHEET_NAME = 'Kết quả';
const HEAD = ['Ngày', 'Giờ', 'Tên nhóm', 'Tên đội', 'Nhiệm vụ', 'Đúng', 'Tổng', 'Điểm (thang 10)', 'Số lần kiểm tra', 'Thời gian làm bài', 'Các lựa chọn', 'Cập nhật lúc', 'Mã', 'Dữ liệu (không sửa)'];
const COL_DURATION = 10, COL_ID = 13, COL_DATA = 14;
const ROOMS = { 1: 'Phòng trực quan', 2: 'Phòng thực nghiệm', 3: 'Nhiệm vụ 2 - Biến hóa hình học' };

function styleHead_(sh) {
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#eef5ff');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    styleHead_(sh);
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 180); sh.setColumnWidth(4, 180); sh.setColumnWidth(COL_DURATION, 130); sh.setColumnWidth(COL_DURATION + 1, 360); sh.setColumnWidth(COL_DATA, 120);
  } else {
    // Làm xong mọi thao tác chèn/đổi cột rồi mới ghi lại hàng tiêu đề một lần
    let changed = false;
    const h4 = String(sh.getRange(1, 4).getValue());
    if (h4 === 'Tên học sinh') {
      // Trang tính cũ: chỉ đổi tiêu đề sang "Tên nhóm" / "Tên đội"
      changed = true;
    } else if (h4 !== HEAD[3]) {
      // Trang tính chưa có cột thứ 4: chèn thêm một cột sau cột 3
      sh.insertColumnAfter(3);
      sh.setColumnWidth(4, 180);
      changed = true;
    }
    if (String(sh.getRange(1, COL_DURATION).getValue()) !== HEAD[COL_DURATION - 1]) {
      // Trang tính chưa có cột "Thời gian làm bài": chèn một cột sau cột "Số lần kiểm tra" (cột 9)
      sh.insertColumnAfter(COL_DURATION - 1);
      sh.setColumnWidth(COL_DURATION, 130);
      sh.setColumnWidth(COL_DURATION + 1, 360);
      changed = true;
    }
    if (changed) styleHead_(sh);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 90 → "1p30s", 45 → "45s", 0/không hợp lệ → ""
function fmtDur_(sec) {
  sec = Math.round(Number(sec));
  if (!Number.isFinite(sec) || sec <= 0) return '';
  const m = Math.floor(sec / 60), s = sec % 60;
  return m ? m + 'p' + (s < 10 ? '0' : '') + s + 's' : s + 's';
}

// Keep only the fields the page uses, with sane sizes; the page shows them escaped.
function clean_(r) {
  if (!r || typeof r !== 'object') return null;
  const str = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
  const room = Number(r.room), ok = Number(r.ok), total = Number(r.total);
  if (!ROOMS[room] || !Number.isInteger(total) || total < 1 || total > 10 || !Number.isInteger(ok) || ok < 0 || ok > total) return null;
  const id = str(r.id, 40), day = str(r.day, 10), group = str(r.group, 80), student = str(r.student, 80);
  if (!id || !group || !/^\d{4}-\d{2}-\d{2}$/.test(day)) return null;
  const answers = r.answers && typeof r.answers === 'object' ? r.answers : {};
  if (JSON.stringify(answers).length > 20000) return null;
  // Thời gian làm bài (giây): lấy trường duration, nếu thiếu thì lấy bản sao trong answers.duration; 0 = không có
  const dur = Number(r.duration != null ? r.duration : answers.duration);
  const duration = Number.isFinite(dur) && dur > 0 ? Math.min(3600, Math.round(dur)) : 0;
  return { id, group, student, room, ok, total, day, at: str(r.at, 30), updatedAt: new Date().toISOString(), tries: Math.max(1, Number(r.tries) | 0), duration, answers };
}

// GET …/exec → every saved result, as the page's records.
function doGet() {
  const sh = getSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return json_([]);
  const list = [];
  sh.getRange(2, COL_DATA, n, 1).getValues().forEach(([cell]) => { try { list.push(JSON.parse(cell)) } catch (e) { } });
  return json_(list);
}

// POST …/exec with one result (JSON as plain text) → add its row, or update it if the machine checked again.
function doPost(e) {
  let rec = null, view = {};
  try { const body = JSON.parse(e.postData.contents); rec = clean_(body); view = body.view || {} } catch (err) { }
  if (!rec) return json_({ error: 'invalid' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(), n = sh.getLastRow() - 1;
    const ids = n > 0 ? sh.getRange(2, COL_ID, n, 1).getValues().map(([v]) => String(v)) : [];
    const i = ids.indexOf(rec.id);
    if (i >= 0) {
      // Keep the first check's date
      try { const old = JSON.parse(sh.getRange(i + 2, COL_DATA).getValue()); rec.at = old.at || rec.at; rec.day = old.day || rec.day } catch (err) { }
    }
    // A leading ' keeps a cell as plain text, so Sheets does not turn it into a date, number or formula
    const text = s => "'" + String(s == null ? '' : s);
    const v = (k, n) => text(String(view[k] == null ? '' : view[k]).slice(0, n));
    const row = [v('date', 10), v('time', 5), text(rec.student), text(rec.group), ROOMS[rec.room], rec.ok, rec.total, Math.round(rec.ok / rec.total * 100) / 10, rec.tries, text(fmtDur_(rec.duration)), v('choices', 2000),
      text(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')), text(rec.id), text(JSON.stringify(rec))];
    if (i >= 0) {
      // A re-check keeps the date and time of the first check
      const first = sh.getRange(i + 2, 1, 1, 2).getDisplayValues()[0];
      if (first[0]) { row[0] = text(first[0]); row[1] = text(first[1]) }
      sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    } else sh.appendRow(row);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}