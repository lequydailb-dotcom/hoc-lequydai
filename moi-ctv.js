/* Lời mời làm cộng tác viên trong trang học: khoá trả phí hiện sau khi xong bài đầu, khoá miễn phí chỉ hiện ở trang hoàn thành. */
(function () {
  var app = document.getElementById('app');
  if (!app || !window.MutationObserver) return;
  var mienPhi = /^\/(tang|ba-ngay)\//.test(location.pathname);
  var the = '<div class="moi-ctv" style="margin:28px 0 8px;background:#0E3B3A;color:#F4EFE4;border-radius:16px;padding:22px 20px;display:flex;flex-wrap:wrap;gap:14px;align-items:center;justify-content:space-between">' +
    '<div style="flex:1;min-width:230px"><div style="font-weight:800;font-size:19px;line-height:1.3;color:#fff">Thấy khoá học có ích? Giới thiệu bạn bè, nhận 30% hoa hồng</div>' +
    '<div style="font-size:15px;color:#CFE3D7;margin-top:6px;line-height:1.5">Bạn gửi một đường dẫn. Ai mua qua đường dẫn đó, bạn nhận 30% số tiền họ trả.</div></div>' +
    '<a href="https://lequydai.com/cong-tac-vien/" target="_blank" rel="noopener" style="background:#FD7800;color:#0E3B3A;font-weight:800;padding:13px 20px;border-radius:10px;text-decoration:none;white-space:nowrap">Lấy đường dẫn của tôi</a></div>';
  function xet() {
    try {
      if (typeof FULL === 'undefined' || !FULL || app.querySelector('.moi-ctv')) return;
      var h = location.hash, xong = h.indexOf('chung-nhan') >= 0, nha = !/bai-/.test(h) && !xong;
      if (xong || (nha && !mienPhi && app.querySelector('.dn.on'))) app.insertAdjacentHTML('beforeend', the);
    } catch (e) {}
  }
  new MutationObserver(xet).observe(app, { childList: true });
  xet();
})();
