/* ===================== CẤP KHOÁ TAY (tặng / bán riêng) ===================== */
var CAP_TAY_TANG_ = false;
function capKhoaTay() {
  const items = Object.keys(PRODUCTS).map(function (k) { return { id: k, name: PRODUCTS[k].name, price: PRODUCTS[k].price || 0, may: !!(PRODUCTS[k].zai || PRODUCTS[k].license) }; });
  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;font-size:14px;color:#1b2a4a;margin:0;padding:4px 6px}' +
    'label{display:block;font-weight:bold;margin:10px 0 4px}' +
    'input[type=text],input[type=email],input[type=tel],input[type=number]{width:100%;box-sizing:border-box;padding:8px;border:1px solid #c8cfdc;border-radius:6px;font-size:14px}' +
    '.k{display:flex;gap:8px;align-items:flex-start;font-weight:normal;margin:4px 0}.k small{color:#667}' +
    '.ly{display:flex;gap:16px}.ly label{font-weight:normal;display:flex;gap:6px;align-items:center;margin:4px 0}' +
    'button{margin-top:14px;width:100%;padding:11px;border:0;border-radius:8px;background:#FD7800;color:#fff;font-size:15px;font-weight:bold;cursor:pointer}' +
    'button:disabled{background:#aaa}#kq{margin-top:10px;white-space:pre-line}.loi{color:#c0392b}.ok{color:#1e8449}' +
    '</style></head><body>' +
    '<label>Họ tên</label><input id="name" type="text">' +
    '<label>Email (học viên đăng nhập bằng email này)</label><input id="email" type="email">' +
    '<label>Số điện thoại</label><input id="phone" type="tel">' +
    '<label>Khoá / sản phẩm cần cấp</label><div id="ds"></div>' +
    '<div id="mayBox" style="display:none"><label>Mã máy của khách (Zalo AI / app cần mã máy)</label><input id="may" type="text" placeholder="Khách mở app, chép mã máy gửi anh"></div>' +
    '<label>Lý do</label><div class="ly"><label><input type="radio" name="ly" value="tang" checked> Tặng (0đ)</label><label><input type="radio" name="ly" value="ban"> Bán riêng</label></div>' +
    '<div id="tienBox" style="display:none"><label>Số tiền đã thu (đồng)</label><input id="tien" type="number" min="0"></div>' +
    '<label>Ghi chú (không bắt buộc)</label><input id="note" type="text" placeholder="Ví dụ: học viên lớp offline tháng 11">' +
    '<button id="nut">Cấp khoá và gửi email</button><div id="kq"></div>' +
    '<script>var SP=' + JSON.stringify(items) + ';' +
    'document.getElementById("ds").innerHTML=SP.map(function(p){return "<label class=k><input type=checkbox value=\\""+p.id+"\\"><span>"+p.name+(p.price?" <small>("+p.price.toLocaleString("vi-VN")+"đ)</small>":" <small>(miễn phí)</small>")+"</span></label>"}).join("");' +
    'function xemMay(){var c=[].slice.call(document.querySelectorAll("#ds input:checked")).some(function(x){return SP.filter(function(p){return p.id===x.value})[0].may});document.getElementById("mayBox").style.display=c?"block":"none"}' +
    'document.querySelectorAll("#ds input").forEach(function(x){x.onchange=xemMay});' +
    'document.querySelectorAll("input[name=ly]").forEach(function(r){r.onchange=function(){document.getElementById("tienBox").style.display=r.value==="ban"&&r.checked?"block":"none"}});' +
    'document.getElementById("nut").onclick=function(){var d={name:name.value,email:email.value,phone:phone.value,note:note.value,may:may.value,' +
    'ly:document.querySelector("input[name=ly]:checked").value,tien:tien.value,ids:[].slice.call(document.querySelectorAll("#ds input:checked")).map(function(x){return x.value})};' +
    'var kq=document.getElementById("kq"),b=this;b.disabled=true;kq.className="";kq.textContent="Đang cấp…";' +
    'google.script.run.withSuccessHandler(function(r){b.disabled=false;kq.className=r.ok?"ok":"loi";kq.textContent=r.msg;if(r.ok){name.value="";email.value="";phone.value="";note.value="";tien.value="";may.value="";document.querySelectorAll("#ds input").forEach(function(x){x.checked=false})}})' +
    '.withFailureHandler(function(e){b.disabled=false;kq.className="loi";kq.textContent="Lỗi: "+e.message}).capKhoaTayGui(d)};' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(640), 'Cấp khoá tay · ĐẠI LÊ AI');
}

function capKhoaTayGui(d) {
  const name = String(d.name || '').trim().slice(0, 80);
  const email = String(d.email || '').trim().toLowerCase().slice(0, 120);
  const phone = String(d.phone || '').replace(/[^\d+]/g, '').slice(0, 15);
  const ids = (d.ids || []).filter(function (k) { return PRODUCTS[k]; });
  const tang = d.ly !== 'ban';
  const tien = tang ? 0 : Math.max(0, Math.round(Number(d.tien) || 0));
  if (!name) return { ok: false, msg: 'Anh điền họ tên giúp em.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, msg: 'Email chưa đúng, anh kiểm tra lại.' };
  if (phone.length < 9) return { ok: false, msg: 'Số điện thoại chưa đúng.' };
  if (!ids.length) return { ok: false, msg: 'Anh tick ít nhất một khoá.' };
  const may = String(d.may || '').trim().slice(0, 80);
  if (!may && ids.some(function (k) { return PRODUCTS[k].zai || PRODUCTS[k].license; })) return { ok: false, msg: 'Zalo AI / app có mã máy: anh dán mã máy của khách vào giúp em.' };
  if (!tang && !tien) return { ok: false, msg: 'Bán riêng thì anh gõ số tiền đã thu.' };
  const note = 'CẤP TAY · ' + (tang ? 'Tặng' : 'Bán riêng') + (d.note ? ' · ' + String(d.note).slice(0, 120) : '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  const codes = [];
  try {
    const sh = sheet_(SH.ORDERS);
    const used = sh.getRange('B:B').getValues().map(function (x) { return String(x[0]); });
    ids.forEach(function (pid, n) {
      let code;
      do { code = 'TAY' + (PRODUCTS[pid].prefix || 'X') + Math.floor(10000 + Math.random() * 90000); } while (used.indexOf(code) >= 0);
      used.push(code); codes.push(code);
      const gia = tang ? 0 : (n === 0 ? tien : 0);
      sh.appendRow([new Date(), code, name, email, "'" + phone, gia, ST.PAID, 'CẤP TAY', '', note, pid, (PRODUCTS[pid].zai || PRODUCTS[pid].license) ? may : '']);
    });
    CAP_TAY_TANG_ = tang;
    provisionPaid_();
    log_('cấp tay', codes.join(',') + ' · ' + email + ' · ' + note);
  } finally { CAP_TAY_TANG_ = false; lock.releaseLock(); }
  return { ok: true, msg: 'Xong. Đã cấp ' + ids.map(function (k) { return PRODUCTS[k].name; }).join(', ') + ' cho ' + email + '.\nMã đơn: ' + codes.join(', ') + '.\nHọc viên nhận email đăng nhập trong vài phút (nhắc kiểm tra cả thư Quảng cáo/Spam).' };
}

