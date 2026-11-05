/**
 * VẬN HÀNH AGENT – Hệ thống bán khoá học tự động · Đại Lê AI
 * Chạy trên Google Apps Script, gắn với Google Sheet "VẬN HÀNH AGENT – Đơn hàng & Học viên".
 *
 * Luồng: khách điền form ở hoc.lequydai.com/dang-ky -> tạo đơn + mail hướng dẫn chuyển khoản
 *        -> khách quét QR chuyển khoản (nội dung có mã đơn VHxxxxx)
 *        -> mỗi phút hệ thống đọc giao dịch SePay (hoặc anh tự chọn "ĐÃ THANH TOÁN" trong Sheet)
 *        -> tự tạo tài khoản + gửi mail đăng nhập -> khách vào hoc.lequydai.com học.
 */

const PRODUCTS = {
  vanhanh: { name: 'VẬN HÀNH AGENT', price: 1990000, prefix: 'VH', file: 'vanhanh-lessons.json', path: '/van-hanh/' },
  vibe: { name: '28 NGÀY VIBE CODING – Bot CSKH Zalo & Facebook', price: 2990000, prefix: 'VC', file: 'vibecoding-lessons.json', path: '/vibe-coding/' },
  donggoi: { name: 'ĐÓNG GÓI KIẾN THỨC THÀNH KHOÁ HỌC BẰNG AI', price: 1490000, prefix: 'DG', file: 'donggoi-lessons.json', path: '/dong-goi/' }
};
const CFG = {
  BANK: 'VietinBank',
  ACC: '104880553933',
  HOLDER: 'LE QUY DAI',
  PREFIX: 'VH',
  SITE: 'https://hoc.lequydai.com',
  COURSE: 'VẬN HÀNH AGENT',
  BRAND: 'Đại Lê AI',
  ZALO_GROUP: 'https://zalo.me/g/a9my3sx1ov0fsihenx6e',
  ZALO_PHONE: '0972 782 782',
  TOKEN_DAYS: 60
};

const ST = { WAIT: 'CHỜ THANH TOÁN', PAID: 'ĐÃ THANH TOÁN', DONE: 'ĐÃ CẤP TÀI KHOẢN', CANCEL: 'HUỶ' };
const SH = { ORDERS: 'DON_HANG', STUDENTS: 'HOC_VIEN', CONFIG: 'CAU_HINH', LOG: 'NHAT_KY' };
const ORDER_HEAD = ['Thời gian', 'Mã đơn', 'Họ tên', 'Email', 'Số điện thoại', 'Số tiền', 'Trạng thái', 'Mã giao dịch', 'Ngày cấp TK', 'Ghi chú', 'Khoá học'];
const STU_HEAD = ['Email', 'Họ tên', 'Số điện thoại', 'Mã đơn', 'Hash mật khẩu', 'Salt', 'Token', 'Token hết hạn', 'Ngày cấp', 'Trạng thái', 'Khoá học'];

/* ===================== CÀI ĐẶT LẦN ĐẦU ===================== */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const mk = (name, head, color) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (head) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const o = mk(SH.ORDERS, ORDER_HEAD, '#0d1838');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList([ST.WAIT, ST.PAID, ST.DONE, ST.CANCEL], true).build();
  o.getRange('G2:G2000').setDataValidation(rule);
  o.getRange('F2:F2000').setNumberFormat('#,##0');
  o.setColumnWidths(1, 11, 140);
  const stu = mk(SH.STUDENTS, STU_HEAD, '#16285a');
  stu.hideColumns(5, 4);
  const c = mk(SH.CONFIG, ['Khoá', 'Giá trị', 'Ghi chú'], '#5d6785');
  if (c.getLastRow() < 2) {
    c.getRange(2, 1, 2, 3).setValues([
      ['SEPAY_API_TOKEN', '', 'Dán API Token SePay vào ô bên trái (my.sepay.vn > Cấu hình công ty > API Access)'],
      ['TU_DONG_DUYET', 'BẬT', 'BẬT = tiền vào là tự cấp tài khoản. TẮT = chỉ cấp khi anh chọn ĐÃ THANH TOÁN']
    ]);
    c.setColumnWidth(2, 380); c.setColumnWidth(3, 520);
  }
  mk(SH.LOG, ['Thời gian', 'Việc', 'Chi tiết'], '#999999');
  const d = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  // Hẹn giờ chạy mỗi phút
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
  log_('setup', 'Đã tạo các trang tính và hẹn giờ mỗi phút');
}

/* ===================== API CHO TRANG WEB ===================== */
function doGet(e) {
  return json_({ ok: true, app: CFG.BRAND, products: Object.keys(PRODUCTS) });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, msg: 'Dữ liệu không hợp lệ' }); }
  try {
    switch (req.action) {
      case 'order': return json_(createOrder_(req));
      case 'status': return json_(orderStatus_(req.code));
      case 'login': return json_(login_(req.email, req.password));
      case 'content': return json_(content_(req.token, req.course));
      case 'me': return json_(me_(req.token));
      case 'progress': return json_(saveProgress_(req.token, req.course));
      case 'quiz': return json_(submitQuiz_(req.token, req.course, req.lesson, req.answers));
      case 'reset': return json_(resetPassword_(req.email));
      default: return json_({ ok: false, msg: 'Không rõ yêu cầu' });
    }
  } catch (err) {
    log_('lỗi', String(err && err.stack || err));
    return json_({ ok: false, msg: 'Hệ thống đang bận, anh/chị thử lại sau ít phút' });
  }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ===================== ĐƠN HÀNG ===================== */
function createOrder_(r) {
  const name = String(r.name || '').trim().slice(0, 80);
  const email = String(r.email || '').trim().toLowerCase().slice(0, 120);
  const phone = String(r.phone || '').replace(/[^\d+]/g, '').slice(0, 15);
  const pid = PRODUCTS[r.course] ? r.course : 'vanhanh';
  const P = PRODUCTS[pid];
  if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || phone.length < 9) return { ok: false, msg: 'Anh/chị điền đủ họ tên, email và số điện thoại đúng giúp em nhé' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.ORDERS);
    // Nếu email này còn đơn đang chờ thì dùng lại, tránh tạo nhiều mã
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][3]).toLowerCase() === email && rows[i][6] === ST.WAIT && (rows[i][10] || 'vanhanh') === pid) return payInfo_(rows[i][1], name, pid);
    }
    let code;
    const used = new Set(rows.map(x => String(x[1])));
    do { code = P.prefix + Math.floor(10000 + Math.random() * 90000); } while (used.has(code));
    sh.appendRow([new Date(), code, name, email, "'" + phone, P.price, ST.WAIT, '', '', '', pid]);
    sendOrderMail_(email, name, code, pid);
    log_('đơn mới', code + ' · ' + pid + ' · ' + email);
    return payInfo_(code, name, pid);
  } finally { lock.releaseLock(); }
}

function payInfo_(code, name, pid) {
  const P = PRODUCTS[pid];
  return {
    ok: true, code: code, name: name, course: pid, amount: P.price, bank: CFG.BANK, acc: CFG.ACC, holder: CFG.HOLDER,
    memo: memo_(code), qr: qrUrl_(code, P.price)
  };
}

/* VietinBank cá nhân qua SePay: nội dung chuyển khoản BẮT BUỘC có tiền tố SEVQR, nếu không SePay không nhận được giao dịch */
function memo_(code) { return 'SEVQR ' + code; }
function qrUrl_(code, amount) {
  return 'https://qr.sepay.vn/img?acc=' + CFG.ACC + '&bank=' + CFG.BANK + '&amount=' + amount + '&des=' + encodeURIComponent(memo_(code)) + '&template=compact';
}
function prefixRe_() {
  return new RegExp('(' + Object.keys(PRODUCTS).map(k => PRODUCTS[k].prefix).join('|') + ')\\d{5}', 'i');
}

function orderStatus_(code) {
  code = String(code || '').toUpperCase();
  const rows = sheet_(SH.ORDERS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][1]) === code) {
    return { ok: true, status: rows[i][6], paid: rows[i][6] === ST.PAID || rows[i][6] === ST.DONE, ready: rows[i][6] === ST.DONE };
  }
  return { ok: false, msg: 'Không tìm thấy đơn' };
}

/* ===================== CHẠY MỖI PHÚT ===================== */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    if (cfg_('TU_DONG_DUYET') !== 'TẮT') checkSepay_();
    provisionPaid_();
    remindDaily_();
  } catch (err) {
    log_('lỗi tick', String(err && err.stack || err));
  } finally { lock.releaseLock(); }
}

function checkSepay_() {
  const token = cfg_('SEPAY_API_TOKEN');
  if (!token) return;
  const from = Utilities.formatDate(new Date(Date.now() - 3 * 86400000), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
  const url = 'https://userapi.sepay.vn/v2/transactions?per_page=100&transaction_date_from=' + encodeURIComponent(from);
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) { log_('SePay', 'Lỗi ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300)); return; }
  const body = JSON.parse(res.getContentText());
  const list = Array.isArray(body.data) ? body.data : (body.data && body.data.data) || body.transactions || [];
  if (!list.length) return;
  const props = PropertiesService.getScriptProperties();
  const seen = new Set(JSON.parse(props.getProperty('SEEN') || '[]'));
  const sh = sheet_(SH.ORDERS);
  const rows = sh.getDataRange().getValues();
  const re = prefixRe_();
  let changed = false;
  list.forEach(t => {
    const id = String(t.id);
    if (seen.has(id)) return;
    seen.add(id); changed = true;
    const amount = Number(t.amount_in || t.transferAmount || 0);
    const acc = String(t.account_number || t.accountNumber || '');
    const text = String(t.transaction_content || t.content || '') + ' ' + String(t.code || '');
    const m = text.match(re);
    if (!m || amount <= 0) return;
    if (acc && acc.indexOf(CFG.ACC) < 0) return;
    const code = m[0].toUpperCase();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === code && rows[i][6] === ST.WAIT) {
        if (amount >= Number(rows[i][5] || 0)) {
          sh.getRange(i + 1, 7, 1, 2).setValues([[ST.PAID, id]]);
          log_('tiền vào', code + ' · ' + amount);
        } else {
          sh.getRange(i + 1, 10).setValue('Chuyển thiếu: ' + amount + ' (GD ' + id + ')');
          log_('chuyển thiếu', code + ' · ' + amount);
        }
      }
    }
  });
  if (changed) props.setProperty('SEEN', JSON.stringify(Array.from(seen).slice(-800)));
}

function provisionPaid_() {
  const sh = sheet_(SH.ORDERS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][6] !== ST.PAID) continue;
    const [, code, name, email, phone] = rows[i];
    const pid = PRODUCTS[rows[i][10]] ? rows[i][10] : 'vanhanh';
    const pw = createAccount_(String(email).toLowerCase(), name, phone, code, pid);
    sendWelcomeMail_(email, name, pw, pid);
    sh.getRange(i + 1, 7).setValue(ST.DONE);
    sh.getRange(i + 1, 9).setValue(new Date());
    log_('cấp tài khoản', code + ' · ' + email);
  }
}

/* ===================== TÀI KHOẢN ===================== */
/* Trả về mật khẩu mới nếu là học viên mới; trả về null nếu email đã có tài khoản (chỉ mở thêm khoá) */
function createAccount_(email, name, phone, code, pid) {
  const sh = sheet_(SH.STUDENTS);
  const i = findStudent_(email);
  if (i > 0) {
    const cell = sh.getRange(i, 11);
    const list = String(cell.getValue() || 'vanhanh').split(',').map(x => x.trim()).filter(Boolean);
    if (list.indexOf(pid) < 0) list.push(pid);
    cell.setValue(list.join(','));
    sh.getRange(i, 4).setValue(String(sh.getRange(i, 4).getValue()) + ',' + code);
    return null;
  }
  const pw = randomPw_();
  const salt = Utilities.getUuid();
  sh.appendRow([email, name, "'" + String(phone).replace(/^'/, ''), code, hash_(salt, pw), salt, '', '', new Date(), 'HOẠT ĐỘNG', pid]);
  return pw;
}

function findStudent_(email) {
  const v = sheet_(SH.STUDENTS).getRange('A:A').getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === email) return i + 1;
  return -1;
}

function login_(email, password) {
  email = String(email || '').trim().toLowerCase();
  const i = findStudent_(email);
  if (i < 0) return { ok: false, msg: 'Email này chưa có tài khoản. Anh/chị kiểm tra lại email đã dùng khi đăng ký.' };
  const sh = sheet_(SH.STUDENTS);
  const r = sh.getRange(i, 1, 1, STU_HEAD.length).getValues()[0];
  if (r[9] !== 'HOẠT ĐỘNG') return { ok: false, msg: 'Tài khoản đang tạm khoá. Anh/chị nhắn Zalo ' + CFG.ZALO_PHONE + ' để được hỗ trợ.' };
  if (hash_(r[5], String(password || '')) !== r[4]) return { ok: false, msg: 'Sai mật khẩu. Anh/chị bấm "Quên mật khẩu" để nhận mật khẩu mới qua email.' };
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  sh.getRange(i, 7, 1, 2).setValues([[token, new Date(Date.now() + CFG.TOKEN_DAYS * 86400000)]]);
  return { ok: true, token: token, name: r[1], courses: String(r[10] || 'vanhanh').split(',') };
}

/* Trang chủ hoc.lequydai.com: biết học viên đang sở hữu khoá nào */
function me_(token) {
  token = String(token || '');
  if (token.length < 30) return { ok: false, login: true };
  const v = sheet_(SH.STUDENTS).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][6] === token) {
      if (v[i][9] !== 'HOẠT ĐỘNG' || new Date(v[i][7]) < new Date()) return { ok: false, login: true };
      return { ok: true, name: v[i][1], courses: String(v[i][10] || 'vanhanh').split(',').map(x => x.trim()) };
    }
  }
  return { ok: false, login: true };
}

function content_(token, course) {
  const pid = PRODUCTS[course] ? course : 'vanhanh';
  token = String(token || '');
  if (token.length < 30) return { ok: false, login: true };
  const v = sheet_(SH.STUDENTS).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][6] === token) {
      if (v[i][9] !== 'HOẠT ĐỘNG' || new Date(v[i][7]) < new Date()) return { ok: false, login: true };
      const owned = String(v[i][10] || 'vanhanh').split(',').map(x => x.trim());
      if (owned.indexOf(pid) < 0) return { ok: false, notOwned: true, name: v[i][1] };
      const L = JSON.parse(JSON.stringify(lessons_(pid)));
      Object.keys(L).forEach(k => { const q = quizPublic_(pid, k); if (q) L[k].quiz = q; });
      touch_(v[i][0], v[i][1], pid);
      const pr = getProgress_(v[i][0], pid);
      return { ok: true, name: v[i][1], courses: owned, lessons: L, progress: { done: pr.done, doneAt: pr.doneAt } };
    }
  }
  return { ok: false, login: true };
}

/* ===================== TIẾN ĐỘ HỌC + TRẮC NGHIỆM + NHẮC HỌC ===================== */
const TOTAL = { vanhanh: 16, vibe: 19, donggoi: 14 };
const PHEAD = ['Email', 'Họ tên', 'Khoá học', 'Bài đã xong', 'Số bài xong', '% hoàn thành', 'Cập nhật', 'Hoàn thành khoá', 'Lần vào học cuối', 'Nhắc lần cuối', 'Số lần nhắc'];
function progressSheet_() {
  let s = sheet_('TIEN_DO');
  if (!s) { s = SpreadsheetApp.getActive().insertSheet('TIEN_DO'); s.setFrozenRows(1); }
  if (s.getRange(1, 11).getValue() !== PHEAD[10]) {
    s.getRange(1, 1, 1, PHEAD.length).setValues([PHEAD]).setFontWeight('bold').setBackground('#0d1838').setFontColor('#ffffff');
  }
  return s;
}
function getProgress_(email, pid) {
  const v = progressSheet_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === String(email).toLowerCase() && v[i][2] === pid) {
      let d = []; try { d = JSON.parse(v[i][3] || '[]'); } catch (e) {}
      return { done: d, doneAt: v[i][7] ? new Date(v[i][7]).toISOString() : '', row: i + 1, vals: v[i] };
    }
  }
  return { done: [], doneAt: '', row: 0, vals: null };
}
function progressRow_(email, name, pid) {
  const p = getProgress_(email, pid);
  if (p.row) return p;
  progressSheet_().appendRow([email, name, pid, '[]', 0, '0%', new Date(), '', new Date(), '', 0]);
  return getProgress_(email, pid);
}
function student_(token, course) {
  const pid = PRODUCTS[course] ? course : 'vanhanh';
  token = String(token || '');
  if (token.length < 30) return { err: { ok: false, login: true } };
  const v = sheet_(SH.STUDENTS).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][6] !== token) continue;
    if (v[i][9] !== 'HOẠT ĐỘNG' || new Date(v[i][7]) < new Date()) return { err: { ok: false, login: true } };
    const owned = String(v[i][10] || 'vanhanh').split(',').map(x => x.trim());
    if (owned.indexOf(pid) < 0) return { err: { ok: false, notOwned: true } };
    return { email: v[i][0], name: v[i][1], pid: pid };
  }
  return { err: { ok: false, login: true } };
}
function touch_(email, name, pid) {
  try { const p = progressRow_(email, name, pid); progressSheet_().getRange(p.row, 9).setValue(new Date()); } catch (e) {}
}
/* Đọc tiến độ (máy chủ là nguồn đúng; học viên chỉ mở bài mới qua trắc nghiệm) */
function saveProgress_(token, course) {
  const st = student_(token, course); if (st.err) return st.err;
  const p = getProgress_(st.email, st.pid);
  return { ok: true, done: p.done, doneAt: p.doneAt };
}
function quiz_(pid, i) {
  const t = PropertiesService.getScriptProperties().getProperty('QUIZ_' + pid + '_' + i);
  return t ? JSON.parse(t) : null;
}
function quizPublic_(pid, i) {
  const q = quiz_(pid, i); return q ? q.map(x => [x[0], x[1]]) : null;
}
function submitQuiz_(token, course, lesson, answers) {
  const st = student_(token, course); if (st.err) return st.err;
  const pid = st.pid, total = TOTAL[pid]; lesson = Number(lesson);
  if (!(lesson >= 1 && lesson <= total)) return { ok: false, msg: 'Bài không hợp lệ' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const p = progressRow_(st.email, st.name, pid);
    if (p.done.indexOf(lesson) >= 0) return { ok: true, pass: true, done: p.done, doneAt: p.doneAt };
    if (lesson > 1 && p.done.indexOf(lesson - 1) < 0) return { ok: false, msg: 'Hoàn thành bài trước để mở bài này.' };
    const q = quiz_(pid, lesson);
    const wrong = [];
    if (q) q.forEach((x, k) => { if (Number((answers || [])[k]) !== x[2]) wrong.push(k); });
    if (wrong.length) { log_('trắc nghiệm chưa đạt', st.email + ' · ' + pid + ' bài ' + lesson); return { ok: true, pass: false, wrong: wrong }; }
    const list = p.done.concat([lesson]).sort((a, b) => a - b);
    const finished = list.length >= total ? (p.doneAt ? new Date(p.doneAt) : new Date()) : '';
    progressSheet_().getRange(p.row, 1, 1, 9).setValues([[st.email, st.name, pid, JSON.stringify(list), list.length, Math.round(list.length / total * 100) + '%', new Date(), finished, new Date()]]);
    if (finished && !p.doneAt) log_('hoàn thành khoá', st.email + ' · ' + pid);
    return { ok: true, pass: true, done: list, doneAt: finished ? finished.toISOString() : '' };
  } finally { lock.releaseLock(); }
}
/* Chạy 1 lần trong trình soạn thảo để nạp bộ câu hỏi (đáp án chỉ nằm trên máy chủ) */
function importQuiz() {
  const url = PropertiesService.getScriptProperties().getProperty('QUIZ_URL') || 'https://raw.githubusercontent.com/lequydailb-dotcom/hoc-lequydai/da183fd2e805dae58427dc32442047b9ccd84ddf/tmp-q.json';
  const data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  const props = PropertiesService.getScriptProperties(); let n = 0;
  Object.keys(data).forEach(pid => Object.keys(data[pid]).forEach(i => { props.setProperty('QUIZ_' + pid + '_' + i, JSON.stringify(data[pid][i])); n++; }));
  props.deleteProperty('QUIZ_URL');
  log_('nạp trắc nghiệm', n + ' bài');
  return n;
}
/* Email nhắc học: mỗi ngày 1 lần (sau 19h), học viên 3 ngày chưa vào học, tối đa 3 lần */
function remindDaily_() {
  const now = new Date(), props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  const hour = Number(Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'H'));
  if (hour < 19 || props.getProperty('REMIND_DATE') === today) return;
  props.setProperty('REMIND_DATE', today);
  if (cfg_('NHAC_HOC') === 'TẮT') return;
  const DAY = 86400000, stu = sheet_(SH.STUDENTS).getDataRange().getValues();
  let sent = 0;
  for (let i = 1; i < stu.length; i++) {
    const email = stu[i][0], name = stu[i][1];
    if (!email || stu[i][9] !== 'HOẠT ĐỘNG') continue;
    String(stu[i][10] || 'vanhanh').split(',').map(x => x.trim()).forEach(pid => {
      if (!PRODUCTS[pid]) return;
      const p = progressRow_(email, name, pid), v = p.vals;
      if (p.doneAt || p.done.length >= TOTAL[pid]) return;
      const last = Math.max(new Date(v[8] || 0).getTime(), new Date(v[6] || 0).getTime(), new Date(stu[i][8] || 0).getTime());
      const lastRem = v[9] ? new Date(v[9]).getTime() : 0, cnt = Number(v[10] || 0);
      if (now - last < 3 * DAY || now - lastRem < 3 * DAY || cnt >= 3) return;
      let next = 1; while (p.done.indexOf(next) >= 0) next++;
      try {
        sendRemindMail_(email, name, pid, next, p.done.length);
        progressSheet_().getRange(p.row, 10, 1, 2).setValues([[now, cnt + 1]]); sent++;
      } catch (e) { log_('lỗi gửi nhắc học', email + ' ' + e); }
    });
  }
  if (sent) log_('gửi nhắc học', sent + ' email');
}
function sendRemindMail_(email, name, pid, next, doneN) {
  let title = '';
  try { const L = lessons_(pid)[next]; title = (L && L.title) || ''; } catch (e) {}
  const P = PRODUCTS[pid], url = CFG.SITE + P.path + '#bai-' + next;
  const lab = pid === 'vibe' ? 'Bài ' + next : 'Bài ' + String(next).padStart(2, '0');
  const first = String(name || '').trim().split(' ').pop() || 'bạn';
  const inner = '<p>Chào ' + esc_(first) + ',</p>' +
    '<p>' + (doneN ? 'Bạn đã đi được <b>' + doneN + '/' + TOTAL[pid] + ' bài</b> của khoá <b>' + P.name + '</b>. Mấy hôm nay chưa thấy bạn vào học, bài tiếp theo đang chờ bạn đây:' : 'Tài khoản khoá <b>' + P.name + '</b> của bạn đã sẵn sàng mà chưa thấy bạn bắt đầu. Bài đầu tiên chỉ mất khoảng 15 phút:') + '</p>' +
    '<p style="font-size:18px;font-weight:700;margin:18px 0">📘 ' + lab + (title ? ': ' + esc_(title) : '') + '</p>' +
    '<p>Mỗi bài có video ngắn khoảng 1 phút và câu lệnh chép sẵn. Học 15 phút mỗi ngày là đủ.</p>' +
    '<p style="margin:24px 0"><a href="' + url + '" style="background:#FFC83D;color:#0b1533;font-weight:800;padding:14px 26px;border-radius:10px;text-decoration:none">VÀO HỌC TIẾP →</a></p>' +
    '<p>Vướng chỗ nào thì nhắn vào nhóm Zalo học viên: <a href="' + CFG.ZALO_GROUP + '">' + CFG.ZALO_GROUP + '</a></p>';
  MailApp.sendEmail({ to: email, subject: (name ? first + ' ơi, ' : '') + lab + ' đang chờ bạn – ' + P.name, htmlBody: mailWrap_(lab + ' đang chờ bạn', inner, pid), name: 'ĐẠI LÊ AI' });
}

function resetPassword_(email) {
  email = String(email || '').trim().toLowerCase();
  const i = findStudent_(email);
  // Luôn trả lời giống nhau để người lạ không dò được email nào đã mua
  if (i > 0) {
    const sh = sheet_(SH.STUDENTS);
    const r = sh.getRange(i, 1, 1, STU_HEAD.length).getValues()[0];
    const pw = randomPw_(), salt = Utilities.getUuid();
    sh.getRange(i, 5, 1, 4).setValues([[hash_(salt, pw), salt, '', '']]);
    sendResetMail_(email, r[1], pw);
    log_('đổi mật khẩu', email);
  }
  return { ok: true, msg: 'Nếu email này đã đăng ký, mật khẩu mới vừa được gửi vào hộp thư (xem cả mục Spam/Quảng cáo).' };
}

function lessons_(pid) {
  const file = PRODUCTS[pid].file;
  const cache = CacheService.getScriptCache();
  const c = cache.get('L_' + pid);
  if (c) return JSON.parse(c);
  const it = DriveApp.getFilesByName(file);
  if (!it.hasNext()) throw new Error('Không tìm thấy file ' + file + ' trên Google Drive');
  const txt = it.next().getBlob().getDataAsString('UTF-8');
  try { cache.put('L_' + pid, txt, 600); } catch (e) {}
  return JSON.parse(txt);
}

/* ===================== EMAIL ===================== */
function mailWrap_(title, inner, pid) {
  const courseName = PRODUCTS[pid || 'vanhanh'].name;
  return '<div style="font-family:Arial,sans-serif;background:#f4f6fb;padding:24px">' +
    '<div style="max-width:560px;margin:auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e3e8f2">' +
    '<div style="background:#0d1838;padding:18px 24px;color:#FFC83D;font-weight:900;font-size:20px">' + courseName + '<div style="color:#c9d4f2;font-size:13px;font-weight:600">' + CFG.BRAND + '</div></div>' +
    '<div style="padding:24px;color:#1b2236;font-size:15px;line-height:1.6"><h2 style="margin:0 0 12px;font-size:19px">' + title + '</h2>' + inner +
    '<p style="margin-top:24px;color:#5d6785;font-size:13px">Cần hỗ trợ: Zalo ' + CFG.ZALO_PHONE + ' · Nhóm học viên: <a href="' + CFG.ZALO_GROUP + '">' + CFG.ZALO_GROUP + '</a></p></div></div></div>';
}

function sendOrderMail_(email, name, code, pid) {
  const P = PRODUCTS[pid];
  const html = mailWrap_('Chào ' + esc_(name) + ', đơn ' + code + ' đã được ghi nhận',
    '<p>Anh/chị chuyển khoản theo thông tin dưới đây. Khi tiền vào, hệ thống tự gửi email đăng nhập trong vài phút.</p>' +
    '<table style="border-collapse:collapse;width:100%;font-size:15px">' +
    row_('Ngân hàng', CFG.BANK) + row_('Số tài khoản', CFG.ACC) + row_('Chủ tài khoản', CFG.HOLDER) +
    row_('Số tiền', fmt_(P.price) + 'đ') + row_('Nội dung', '<b style="color:#D32F2F;font-size:18px">' + memo_(code) + '</b>') + '</table>' +
    '<p style="text-align:center"><img src="' + qrUrl_(code, P.price) + '" alt="QR chuyển khoản" style="width:240px;max-width:100%"></p>' +
    '<p><b>Lưu ý:</b> ghi đúng nội dung <b>' + memo_(code) + '</b> (có cả chữ SEVQR) để hệ thống nhận ra đơn của anh/chị.</p>', pid);
  MailApp.sendEmail({ to: email, subject: '[' + P.name + '] Hướng dẫn chuyển khoản – đơn ' + code, htmlBody: html, name: CFG.BRAND });
}

function sendWelcomeMail_(email, name, pw, pid) {
  const P = PRODUCTS[pid || 'vanhanh'];
  const url = CFG.SITE + P.path;
  const acc = pw
    ? row_('Email đăng nhập', esc_(email)) + row_('Mật khẩu', '<b style="font-size:18px;letter-spacing:1px">' + pw + '</b>')
    : row_('Email đăng nhập', esc_(email)) + row_('Mật khẩu', 'Dùng mật khẩu cũ của anh/chị (quên thì bấm "Quên mật khẩu")');
  const html = mailWrap_('Chào mừng ' + esc_(name) + ' vào khoá ' + P.name + ' 🎉',
    '<p>Đã nhận được thanh toán. ' + (pw ? 'Tài khoản học của anh/chị đã sẵn sàng:' : 'Khoá học mới đã được mở trong tài khoản hiện có của anh/chị:') + '</p>' +
    '<table style="border-collapse:collapse;width:100%;font-size:15px">' +
    row_('Trang học', '<a href="' + url + '">' + url.replace('https://', '') + '</a>') + acc + '</table>' +
    '<p style="text-align:center;margin:22px 0"><a href="' + url + '" style="background:#FFC83D;color:#0d1838;font-weight:900;padding:14px 28px;border-radius:10px;text-decoration:none;display:inline-block">VÀO HỌC NGAY</a></p>' +
    '<p>Vào nhóm Zalo học viên để hỏi đáp và nhận cập nhật: <a href="' + CFG.ZALO_GROUP + '">' + CFG.ZALO_GROUP + '</a></p>' +
    '<p style="color:#5d6785;font-size:13px">Anh/chị giữ email này. Quên mật khẩu thì bấm "Quên mật khẩu" ở trang đăng nhập.</p>', pid);
  MailApp.sendEmail({ to: email, subject: '[' + P.name + '] ' + (pw ? 'Tài khoản học của anh/chị đã sẵn sàng' : 'Đã mở khoá học mới'), htmlBody: html, name: CFG.BRAND });
}

function sendResetMail_(email, name, pw) {
  const html = mailWrap_('Mật khẩu mới của ' + esc_(name),
    '<p>Mật khẩu mới để vào <a href="' + CFG.SITE + '">' + CFG.SITE.replace('https://', '') + '</a> (dùng chung cho mọi khoá đã mua):</p>' +
    '<p style="font-size:22px;font-weight:900;letter-spacing:1px">' + pw + '</p>' +
    '<p style="color:#5d6785;font-size:13px">Nếu anh/chị không yêu cầu đổi mật khẩu, hãy nhắn Zalo ' + CFG.ZALO_PHONE + '.</p>');
  MailApp.sendEmail({ to: email, subject: '[' + CFG.BRAND + '] Mật khẩu mới', htmlBody: html, name: CFG.BRAND });
}

/* ===================== TIỆN ÍCH ===================== */
function sheet_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function cfg_(key) {
  const v = sheet_(SH.CONFIG).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) return String(v[i][1]).trim();
  return '';
}
function hash_(salt, pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function randomPw_() {
  const a = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += a[Math.floor(Math.random() * a.length)];
  return s;
}
function fmt_(n) { return Number(n).toLocaleString('vi-VN'); }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function row_(k, v) { return '<tr><td style="padding:8px;border:1px solid #e3e8f2;color:#5d6785;width:40%">' + k + '</td><td style="padding:8px;border:1px solid #e3e8f2">' + v + '</td></tr>'; }
function log_(what, detail) {
  try { const s = sheet_(SH.LOG); if (s) s.appendRow([new Date(), what, detail]); } catch (e) {}
}

/* Chạy tay để thử: tạo 1 đơn thử với email của anh */
function testOrder() {
  Logger.log(JSON.stringify(createOrder_({ name: 'Đơn thử', email: Session.getActiveUser().getEmail(), phone: '0972782782' })));
}
