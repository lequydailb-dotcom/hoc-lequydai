/* Ghi nhớ mã giới thiệu 30 ngày và gắn mã vào đơn khi khách đặt mua. */
(function () {
  var API = 'https://script.google.com/macros/s/AKfycbwzrKVxemleL-0ToFb08F9hwl7_8p2f8TmDyr0IXEhhlbSAfyFNaSKBPKaxtNp2d4wo/exec';
  function doc() { var m = document.cookie.match(/(?:^|; )dlai_gt=([A-Z0-9]{3,12})/); return m ? m[1] : ''; }
  var q = '';
  try { q = (new URLSearchParams(location.search).get('gt') || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 12); } catch (e) {}
  if (q.length >= 3) {
    var mien = /(^|\.)lequydai\.com$/.test(location.hostname) ? ';domain=.lequydai.com' : '';
    document.cookie = 'dlai_gt=' + q + ';path=/;max-age=' + 30 * 86400 + mien + ';SameSite=Lax';
    var daDem = false;
    try { daDem = sessionStorage.getItem('dlai_gt_dem') === q; sessionStorage.setItem('dlai_gt_dem', q); } catch (e) {}
    if (!daDem) {
      try { fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'ctvclick', code: q }) }).catch(function () {}); } catch (e) {}
    }
  }
  var goc = window.fetch;
  if (!goc) return;
  window.fetch = function (u, o) {
    try {
      if (o && typeof o.body === 'string' && String(u).indexOf('script.google.com/macros') > 0) {
        var b = JSON.parse(o.body);
        if (/^(order|licorder|zaiorder)$/.test(b.action) && !b.gt) {
          var g = doc();
          if (g) { b.gt = g; o = Object.assign({}, o, { body: JSON.stringify(b) }); }
        }
      }
    } catch (e) {}
    return goc.call(this, u, o);
  };
})();
