/* Nạp 4 khoá mới (15-11-2026): VIẾT HAY BÁN ĐƯỢC, TỰ ĐỘNG VIẾT BÀI, EDIT VIDEO TỰ ĐỘNG, WEBSITE BÁN HÀNG TỰ ĐỘNG */
function importKhoaK4() {
  const base = 'https://raw.githubusercontent.com/lequydailb-dotcom/hoc-lequydai/K4HASH/';
  const props = PropertiesService.getScriptProperties();
  const out = [];
  ['viethay', 'vietbai', 'editvideo', 'website'].forEach(function (pid) {
    const L = JSON.parse(UrlFetchApp.fetch(base + 'tmp-' + pid + '-lessons.json').getContentText('UTF-8'));
    Object.keys(L).forEach(function (i) { props.setProperty('LES_' + pid + '_' + i, JSON.stringify(L[i])); });
    props.setProperty('LES_' + pid + '_keys', JSON.stringify(Object.keys(L)));
    CacheService.getScriptCache().remove('L_' + pid);
    const q = JSON.parse(UrlFetchApp.fetch(base + 'tmp-' + pid + '-quiz.json').getContentText('UTF-8'));
    Object.keys(q[pid]).forEach(function (i) { props.setProperty('QUIZ_' + pid + '_' + i, JSON.stringify(q[pid][i])); });
    out.push(pid + ':' + Object.keys(L).length + '/' + Object.keys(q[pid]).length);
  });
  log_('nạp 4 khoá mới', out.join(' '));
  Logger.log('KIEM ' + out.join(' '));
}

