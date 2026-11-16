/* Khung "vừa đăng ký": chỉ hiện đăng ký THẬT lấy từ máy chủ, chỉ tên gọi, không email, không số điện thoại */
(function(){
var API=window.API_URL||'https://script.google.com/macros/s/AKfycbwzrKVxemleL-0ToFb08F9hwl7_8p2f8TmDyr0IXEhhlbSAfyFNaSKBPKaxtNp2d4wo/exec';
var K='dk-moi',TAT='dk-tat';
function ss(k,v){try{if(v===undefined)return sessionStorage.getItem(k);sessionStorage.setItem(k,v)}catch(e){return null}}
if(ss(TAT))return;
function truoc(t){var p=Math.max(1,Math.round((Date.now()-t)/60000));if(p<60)return p+' phút trước';var g=Math.round(p/60);if(g<24)return g+' giờ trước';return Math.round(g/24)+' ngày trước'}
function dung(ds){if(!ds||!ds.length)return;
 var st=document.createElement('style');st.textContent='#dkm{position:fixed;left:16px;bottom:18px;z-index:60;max-width:330px;background:#fff;border:1px solid #E0D8C6;border-left:5px solid #FD7800;border-radius:12px;box-shadow:0 10px 30px rgba(14,59,58,.18);padding:11px 34px 11px 13px;font:14.5px/1.4 "Be Vietnam Pro",system-ui,Arial,sans-serif;color:#1D3432;transform:translateY(160%);opacity:0;transition:transform .45s,opacity .45s;pointer-events:none}#dkm.hien{transform:none;opacity:1;pointer-events:auto}#dkm b{color:#0E3B3A}#dkm small{display:block;color:#4E6461;font-size:12.5px;margin-top:2px}#dkm button{position:absolute;top:4px;right:6px;border:0;background:none;font-size:18px;color:#4E6461;cursor:pointer;padding:4px;line-height:1}@media(max-width:768px){#dkm{left:10px;right:10px;max-width:none;bottom:86px}}';
 document.head.appendChild(st);
 var o=document.createElement('div');o.id='dkm';o.setAttribute('role','status');o.innerHTML='<span id="dkmn"></span><small id="dkmt"></small><button type="button" aria-label="Tắt thông báo">×</button>';document.body.appendChild(o);
 var i=0,dem=0,hen;
 o.querySelector('button').onclick=function(){o.classList.remove('hien');clearTimeout(hen);ss(TAT,'1')};
 function hien(){if(dem>=Math.min(ds.length,8))return;var x=ds[i%ds.length];i++;dem++;
  var n=document.getElementById('dkmn');n.textContent='';var b=document.createElement('b');b.textContent='Bạn '+x.n;n.appendChild(b);n.appendChild(document.createTextNode(' vừa đăng ký '));var c=document.createElement('b');c.textContent=x.c;n.appendChild(c);
  document.getElementById('dkmt').textContent=truoc(x.t)+' · đăng ký thật';
  o.classList.add('hien');setTimeout(function(){o.classList.remove('hien')},5500);hen=setTimeout(hien,25000)}
 hen=setTimeout(hien,8000)}
var luu=null;try{luu=JSON.parse(ss(K)||'null')}catch(e){}
if(luu&&Date.now()-luu.at<600000)return dung(luu.ds);
fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'recent'})}).then(function(r){return r.json()}).then(function(d){if(d&&d.ok&&d.list){ss(K,JSON.stringify({at:Date.now(),ds:d.list}));dung(d.list)}}).catch(function(){});
})();
