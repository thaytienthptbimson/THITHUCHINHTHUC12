/**
 * THỐNG KÊ KẾT QUẢ HỌC SINH - CHỈ CẦN 1 FILE NÀY
 *
 * Link web app: https://script.google.com/macros/s/AKfycbzdN-FjMwdw6rbIP6a6ux9l7jcvUn1u1Dm5E6AT1mplfujs5gKpy5Ol5_HFOlbH9_zj/exec
 *
 * - Mở link trên bằng trình duyệt  -> trang thống kê cho giáo viên (nhập mật khẩu VIEW_KEY)
 * - Học sinh ở các máy khác nộp kết quả về sheet "KetQua" bằng 1 trong 2 cách:
 *
 *   Cách 1 - từ trang web bất kỳ (GitHub Pages...):
 *     fetch("<link /exec ở trên>", {
 *       method: "POST",
 *       headers: { "Content-Type": "text/plain;charset=utf-8" },
 *       body: JSON.stringify({ hoTen: "Nguyễn An", lop: "9A", bai: "Bài 1",
 *                              diem: 8, tongDiem: 10, thoiGianLam: 300, maMay: "may05" })
 *     });
 *
 *   Cách 2 - nếu trang bài tập cũng chạy trong Apps Script:
 *     google.script.run.nopKetQua({ hoTen: "Nguyễn An", lop: "9A", ... });
 *
 * SAU KHI SỬA CODE: Triển khai → Quản lý bản triển khai → Phiên bản mới → Triển khai.
 */

const CONFIG = {
  SHEET_ID: "",         // Để trống nếu script tạo từ Tiện ích mở rộng → Apps Script trong chính file Sheet
  SHEET_NAME: "KetQua", // Tên sheet lưu kết quả (tự tạo nếu chưa có)
  VIEW_KEY: ""          // Mật khẩu xem thống kê. NÊN đặt. Để trống = ai có link cũng xem được
};

const HEADERS = ["ThoiGian", "HoTen", "Lop", "Bai", "Diem", "TongDiem", "ThoiGianLam", "MaMay"];

function sheet_() {
  const ss = CONFIG.SHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- GHI KẾT QUẢ ----------
function ghi_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const text = v => String(v == null ? "" : v).trim().slice(0, 100);
    const num = v => (isNaN(Number(v)) ? 0 : Number(v));
    if (!text(d.hoTen)) throw new Error("Thiếu họ tên");
    sheet_().appendRow([
      new Date(), text(d.hoTen), text(d.lop), text(d.bai),
      num(d.diem), num(d.tongDiem), num(d.thoiGianLam), text(d.maMay)
    ]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function nopKetQua(d) { // gọi bằng google.script.run
  try { return ghi_(d); } catch (err) { return { ok: false, error: String(err) }; }
}

function doPost(e) { // gọi bằng fetch từ máy khác
  try { return json_(ghi_(JSON.parse(e.postData.contents))); }
  catch (err) { return json_({ ok: false, error: String(err) }); }
}

// ---------- ĐỌC KẾT QUẢ ----------
function layDuLieu(key) {
  try {
    if (CONFIG.VIEW_KEY && key !== CONFIG.VIEW_KEY) return { ok: false, error: "Sai mật khẩu" };
    const rows = sheet_().getDataRange().getValues().slice(1).map(r => ({
      thoiGian: r[0] instanceof Date ? r[0].getTime() : 0,
      hoTen: r[1], lop: r[2], bai: r[3], diem: r[4], tongDiem: r[5], thoiGianLam: r[6], maMay: r[7]
    }));
    return { ok: true, rows: rows };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "data") return json_(layDuLieu(p.key)); // dữ liệu JSON: ?action=data&key=...
  return HtmlService.createHtmlOutput(TRANG_THONG_KE)
    .setTitle("Thống kê kết quả học sinh")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- TRANG THỐNG KÊ (HTML nằm ngay trong file này) ----------
const TRANG_THONG_KE = `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;600;700&display=swap">
<style>
:root{--bg:#f2f5f7;--ink:#14263a;--mute:#5d6b7a;--line:#d5dde4;--card:#fff;--acc:#1f5fbf}
@media (prefers-color-scheme:dark){:root{--bg:#0f1a26;--ink:#e6edf3;--mute:#93a3b3;--line:#263647;--card:#16232f;--acc:#6aa4ff}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 "Be Vietnam Pro",system-ui,sans-serif}
main{max-width:980px;margin:0 auto;padding:24px 16px 48px}
h1{font-size:24px;margin:0 0 16px}
h2{font-size:16px;margin:32px 0 8px}
.top,.stats{display:flex;flex-wrap:wrap;gap:12px 32px;align-items:end}
label{display:grid;gap:4px;font-size:13px;color:var(--mute)}
input,select,button{font:inherit;padding:7px 10px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--ink)}
button{background:var(--acc);border-color:var(--acc);color:#fff;font-weight:600;cursor:pointer}
:focus-visible{outline:2px solid var(--acc);outline-offset:2px}
.stats{margin-top:24px}
.stats b{display:block;font-size:32px;line-height:1.1}
.stats span{color:var(--mute);font-size:13px}
.wrap{overflow-x:auto;background:var(--card);border:1px solid var(--line);border-radius:8px}
table{width:100%;border-collapse:collapse;min-width:560px}
th,td{padding:9px 12px;text-align:left;border-bottom:1px solid var(--line);white-space:nowrap}
th{font-size:13px;color:var(--mute);font-weight:600}
tr:last-child td{border-bottom:0}
.bar{display:inline-block;width:90px;height:8px;border-radius:4px;background:var(--line);vertical-align:middle;margin-right:6px}
.bar i{display:block;height:100%;border-radius:4px;background:var(--acc)}
#msg{color:var(--mute);font-size:13px;margin-top:8px}
</style>
</head>
<body>
<main>
  <h1>Kết quả học sinh</h1>
  <div class="top">
    <label>Mật khẩu xem<input id="key" type="password" autocomplete="off"></label>
    <label>Lớp<select id="lop"></select></label>
    <label>Bài<select id="bai"></select></label>
    <button id="reload">Tải lại</button>
  </div>
  <div id="msg"></div>
  <div class="stats" id="stats"></div>
  <h2>Theo học sinh</h2>
  <div class="wrap"><table>
    <thead><tr><th>Họ tên</th><th>Lớp</th><th>Số lần làm</th><th>Điểm trung bình</th><th>Cao nhất</th><th>Lần cuối</th></tr></thead>
    <tbody id="hs"></tbody>
  </table></div>
  <h2>10 lượt nộp gần nhất</h2>
  <div class="wrap"><table>
    <thead><tr><th>Thời gian</th><th>Họ tên</th><th>Bài</th><th>Điểm</th><th>Mã máy</th></tr></thead>
    <tbody id="gan"></tbody>
  </table></div>
</main>
<script>
var $=function(id){return document.getElementById(id)};
var esc=function(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")};
var fmt=function(t){return new Date(t).toLocaleString("vi-VN",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})};
var all=[];
try{$("key").value=localStorage.getItem("tk_key")||""}catch(e){}

function fill(id,vals){
  var cur=$(id).value,u=[];
  vals.forEach(function(v){if(v&&u.indexOf(v)<0)u.push(v)});
  u.sort();
  $(id).innerHTML='<option value="">Tất cả</option>'+u.map(function(v){return "<option>"+esc(v)+"</option>"}).join("");
  $(id).value=cur;
}

function load(){
  $("msg").textContent="Đang tải…";
  try{localStorage.setItem("tk_key",$("key").value)}catch(e){}
  google.script.run
    .withSuccessHandler(function(d){
      if(!d.ok){$("msg").textContent="Không tải được: "+d.error;return}
      all=d.rows.map(function(r){r.lop=String(r.lop);r.bai=String(r.bai);return r});
      fill("lop",all.map(function(r){return r.lop}));
      fill("bai",all.map(function(r){return r.bai}));
      $("msg").textContent="Cập nhật lúc "+fmt(Date.now());
      render();
    })
    .withFailureHandler(function(e){$("msg").textContent="Không tải được: "+e.message})
    .layDuLieu($("key").value);
}

function render(){
  var lop=$("lop").value,bai=$("bai").value;
  var rows=all.filter(function(r){return (!lop||r.lop===lop)&&(!bai||r.bai===bai)}).map(function(r){
    r.pct=r.tongDiem>0?r.diem/r.tongDiem*100:0;return r});
  var g={},may={};
  rows.forEach(function(r){
    var k=r.hoTen+"|"+r.lop;
    var s=g[k]=g[k]||{hoTen:r.hoTen,lop:r.lop,n:0,sum:0,max:0,last:0};
    s.n++;s.sum+=r.pct;s.max=Math.max(s.max,r.pct);s.last=Math.max(s.last,r.thoiGian);
    may[r.maMay]=1;
  });
  var st=Object.keys(g).map(function(k){return g[k]}).sort(function(a,b){return b.sum/b.n-a.sum/a.n});
  var avg=rows.length?rows.reduce(function(a,r){return a+r.pct},0)/rows.length:0;

  $("stats").innerHTML=[[rows.length,"lượt nộp bài"],[st.length,"học sinh"],[avg.toFixed(0)+"%","điểm trung bình"],[Object.keys(may).length,"máy tính"]]
    .map(function(x){return "<div><b>"+x[0]+"</b><span>"+x[1]+"</span></div>"}).join("");

  $("hs").innerHTML=st.map(function(s){
    var a=s.sum/s.n;
    return "<tr><td>"+esc(s.hoTen)+"</td><td>"+esc(s.lop)+"</td><td>"+s.n+'</td><td><span class="bar"><i style="width:'+Math.min(100,a)+'%"></i></span>'+a.toFixed(0)+"%</td><td>"+s.max.toFixed(0)+"%</td><td>"+fmt(s.last)+"</td></tr>";
  }).join("")||'<tr><td colspan="6">Chưa có kết quả nào.</td></tr>';

  $("gan").innerHTML=rows.slice().sort(function(a,b){return b.thoiGian-a.thoiGian}).slice(0,10).map(function(r){
    return "<tr><td>"+fmt(r.thoiGian)+"</td><td>"+esc(r.hoTen)+"</td><td>"+esc(r.bai)+"</td><td>"+esc(r.diem)+"/"+esc(r.tongDiem)+"</td><td>"+esc(String(r.maMay).slice(0,8))+"</td></tr>";
  }).join("")||'<tr><td colspan="5">Chưa có kết quả nào.</td></tr>';
}

$("reload").onclick=load;
$("lop").onchange=$("bai").onchange=render;
load();
setInterval(load,30000);
</script>
</body>
</html>`;
