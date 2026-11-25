/* Kho câu lệnh của khoá: gom mọi câu lệnh mẫu của các bài đã mở vào một chỗ, bấm là chép. Chỉ chạy cho học viên đã đăng nhập. */
(function () {
  var app = document.getElementById('app');
  if (!app || !window.MutationObserver) return;
  var css = document.createElement('style');
  css.textContent = '.kho-the{margin:26px 0 8px;background:#fff;border:2px solid #0E3B3A;border-radius:16px;padding:20px}' +
    '.kho-the h3{margin:0 0 4px;font-size:20px;color:#0E3B3A}.kho-the p{margin:0 0 12px;color:#4E6461;font-size:15px;line-height:1.5}' +
    '.kho-nut{background:#FD7800;color:#0E3B3A;font-weight:800;border:0;border-radius:10px;padding:12px 18px;font-size:16px;cursor:pointer;font-family:inherit}' +
    '.kho-ds{display:none;margin-top:16px}.kho-ds.mo{display:block}' +
    '.kho-muc{border-top:1px solid #CFE3D7;padding:14px 0}.kho-muc b{display:block;color:#0E3B3A;font-size:15.5px;margin-bottom:8px;line-height:1.4}' +
    '.kho-muc small{color:#B4510A;font-weight:700;font-size:13px;display:block;margin-bottom:2px}' +
    '.kho-muc pre{white-space:pre-wrap;word-break:break-word;background:#0E3B3A;color:#F4EFE4;border-radius:10px;padding:14px;font-size:14.5px;line-height:1.55;margin:0 0 8px;font-family:inherit;position:static}' +
    '.kho-chep{background:transparent;border:2px solid #0E3B3A;color:#0E3B3A;font-weight:700;border-radius:9px;padding:7px 14px;font-size:14.5px;cursor:pointer;font-family:inherit}' +
    '.kho-khoa{color:#4E6461;font-size:14.5px;border-top:1px solid #CFE3D7;padding-top:12px}';
  document.head.appendChild(css);
  function chep(t, nut) {
    function xong() { var c = nut.textContent; nut.textContent = 'Đã chép'; setTimeout(function () { nut.textContent = c; }, 1500); }
    function cu() { var a = document.createElement('textarea'); a.value = t; document.body.appendChild(a); a.select(); try { document.execCommand('copy'); } catch (e) {} a.remove(); xong(); }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(xong, cu); else cu();
  }
  function gom() {
    var ra = [], khoa = 0, n = typeof NL === 'number' ? NL : 0;
    for (var i = 0; i <= n; i++) {
      var L = FULL[i]; if (!L || !L.html) continue;
      var d = document.createElement('div'); d.innerHTML = L.html;
      var pres = d.querySelectorAll('pre'); if (!pres.length) continue;
      var mo = i === 0 || typeof unlocked !== 'function' || unlocked(i);
      if (!mo) { khoa += pres.length; continue; }
      for (var j = 0; j < pres.length; j++) {
        var h = pres[j].previousElementSibling, ten = '';
        while (h && !/^H[2-4]$/.test(h.tagName) && !(h.tagName === 'P' && h.querySelector('b'))) h = h.previousElementSibling;
        if (h) ten = h.textContent.replace(/\(dùng với[^)]*\)/, '').replace(/^Câu lệnh mẫu:\s*/, '').trim();
        ra.push({ bai: i, ten: ten || 'Câu lệnh mẫu', nd: pres[j].textContent });
      }
    }
    return { ds: ra, khoa: khoa };
  }
  function xet() {
    try {
      if (typeof FULL === 'undefined' || !FULL || app.querySelector('.kho-the')) return;
      var h = location.hash; if (/bai-|chung-nhan/.test(h)) return;
      var k = gom(); if (!k.ds.length && !k.khoa) return;
      var the = document.createElement('div'); the.className = 'kho-the';
      the.innerHTML = '<h3>📋 Kho câu lệnh của khoá</h3><p>' + (k.ds.length + k.khoa) + ' câu lệnh mẫu của cả khoá gom về một chỗ. Bấm là chép, dán vào ChatGPT hoặc Claude là chạy.' + (k.khoa ? ' Bạn đã mở ' + k.ds.length + ' câu, ' + k.khoa + ' câu còn lại mở dần khi học tới.' : '') + '</p>' +
        (k.ds.length ? '<button class="kho-nut" type="button">Mở kho câu lệnh (' + k.ds.length + ')</button><div class="kho-ds"></div>' : '');
      var ds = the.querySelector('.kho-ds');
      if (ds) {
        k.ds.forEach(function (x) {
          var m = document.createElement('div'); m.className = 'kho-muc';
          var s = document.createElement('small'); s.textContent = 'Bài ' + (x.bai < 10 ? '0' : '') + x.bai; m.appendChild(s);
          var b = document.createElement('b'); b.textContent = x.ten; m.appendChild(b);
          var p = document.createElement('pre'); p.textContent = x.nd; m.appendChild(p);
          var nut = document.createElement('button'); nut.className = 'kho-chep'; nut.type = 'button'; nut.textContent = 'Chép câu lệnh';
          nut.onclick = function () { chep(x.nd, nut); }; m.appendChild(nut); ds.appendChild(m);
        });
        if (k.khoa) { var e = document.createElement('div'); e.className = 'kho-khoa'; e.textContent = '🔒 Còn ' + k.khoa + ' câu lệnh ở các bài bạn chưa học tới.'; ds.appendChild(e); }
        var nutMo = the.querySelector('.kho-nut');
        nutMo.onclick = function () { var mo = ds.classList.toggle('mo'); nutMo.textContent = mo ? 'Đóng kho câu lệnh' : 'Mở kho câu lệnh (' + k.ds.length + ')'; };
      }
      var moi = app.querySelector('.moi-ctv');
      if (moi) app.insertBefore(the, moi); else app.appendChild(the);
    } catch (e) {}
  }
  new MutationObserver(xet).observe(app, { childList: true });
  xet();
})();
