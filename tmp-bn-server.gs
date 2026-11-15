/* Khoá tặng 3 NGÀY: email nối tiếp ngày 1, 2, 4 */
function sendBaNgayFollow_(email, name, step, doneN) {
  const P = PRODUCTS.bangay, url = CFG.SITE + P.path, first = String(name || '').trim().split(' ').pop() || 'bạn';
  const btn = (href, t) => '<p style="margin:20px 0"><a href="' + href + '" style="background:#FFC83D;color:#0b1533;font-weight:800;padding:13px 24px;border-radius:10px;text-decoration:none">' + t + '</a></p>';
  let subject, inner;
  if (step === 1) {
    subject = first + ' ơi, ngày 2: trang web đầu tiên của bạn';
    inner = '<p>Chào ' + esc_(first) + ',</p><p>' + (doneN >= 2 ? 'Bạn đã xong Ngày 1, có trợ lý AI riêng rồi. ' : 'Ngày 1 vẫn đang chờ bạn: giao việc cho AI và lập trợ lý riêng, khoảng 50 phút. ') +
      'Hôm nay là <b>Ngày 2</b>: mô tả cho AI là có trang web giới thiệu đẹp trên điện thoại, rồi đưa lên mạng để có link gửi khách. Không cần biết code.</p>' + btn(url, 'HỌC TIẾP →');
  } else if (step === 2) {
    subject = 'Ngày 3: món quà giúp ' + first + ' có khách quan tâm';
    inner = '<p>Chào ' + esc_(first) + ',</p><p>Hôm nay là <b>Ngày 3</b>, ngày cuối: gói một mẹo bạn rành nhất thành quà tặng một trang, rồi dùng nó để có người quan tâm thật bằng bài đăng có từ khoá. Bài cuối có tặng bạn <b>30 câu lệnh chép sẵn</b> cho người bán hàng.</p>' +
      (doneN ? '<p>Bạn đã xong ' + doneN + '/6 bài. Còn một chút nữa thôi!</p>' : '') + btn(url, 'HỌC TIẾP →');
  } else {
    subject = 'Sau 3 ngày: bước tiếp theo của ' + first;
    inner = '<p>Chào ' + esc_(first) + ',</p><p>' + (doneN >= TOTAL.bangay ? 'Chúc mừng bạn đã đi hết 3 ngày! Bạn đã làm được bản đầu tiên của cả ba việc.' : 'Bạn đã xong ' + doneN + '/6 bài. Khoá tặng vẫn mở, học tiếp lúc nào cũng được.') + ' Muốn làm sâu hơn, chọn đúng việc bạn cần:</p><ul>' +
      '<li><b>Đội trợ lý AI làm việc thay mình mỗi ngày</b> → <a href="' + CFG.SITE + '/dang-ky/">VẬN HÀNH AGENT</a></li>' +
      '<li><b>Bot tự trả lời khách Zalo, Facebook, không cần code</b> → <a href="' + CFG.SITE + '/vibe-coding/dang-ky/">28 NGÀY VIBE CODING</a></li>' +
      '<li><b>Biến kinh nghiệm thành khoá học và bán tự động</b> → <a href="' + CFG.SITE + '/dong-goi/dang-ky/">ĐÓNG GÓI KIẾN THỨC</a></li></ul>' +
      btn(CFG.SITE, 'XEM CÁC KHOÁ HỌC →');
  }
  inner += '<p>Nhóm Zalo hỏi đáp: <a href="' + CFG.ZALO_GROUP + '">' + CFG.ZALO_GROUP + '</a></p>';
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: mailWrap_(subject, inner, 'bangay'), name: CFG.BRAND });
}

function importKhoaBaNgay() {
  const base = 'https://raw.githubusercontent.com/lequydailb-dotcom/hoc-lequydai/1062e05/';
  const L = JSON.parse(UrlFetchApp.fetch(base + 'tmp-bn-lessons.json').getContentText('UTF-8'));
  const props = PropertiesService.getScriptProperties();
  Object.keys(L).forEach(function (i) { props.setProperty('LES_bangay_' + i, JSON.stringify(L[i])); });
  props.setProperty('LES_bangay_keys', JSON.stringify(Object.keys(L)));
  CacheService.getScriptCache().remove('L_bangay');
  const q = JSON.parse(UrlFetchApp.fetch(base + 'tmp-bn-quiz.json').getContentText('UTF-8'));
  Object.keys(q.bangay).forEach(function (i) { props.setProperty('QUIZ_bangay_' + i, JSON.stringify(q.bangay[i])); });
  log_('nạp khoá tặng 3 ngày', Object.keys(L).length + ' bài');
  Logger.log('KIEM bai=' + Object.keys(L).length + ' quiz=' + Object.keys(q.bangay).length);
}

