/**
 * 경영전략 파트너 연락처 대시보드 - Apps Script
 *
 * ===== 설치 방법 =====
 * 1. 스프레드시트 > [확장 프로그램] > [Apps Script]
 * 2. 기존 코드 전부 삭제 후 이 코드 붙여넣기
 * 3. Ctrl+S 저장
 * 4. [배포] > [새 배포]
 *    - 유형 선택: 웹 앱
 *    - 실행 주체: 본인
 *    - 액세스 권한: 모든 사용자
 * 5. [배포] 클릭 → URL 복사 → 이 URL이 대시보드 주소!
 *
 * ===== 수정 후 재배포 =====
 * [배포] > [배포 관리] > 연필 아이콘 > 버전: 새 버전 > [배포]
 */

// ===== 설정 =====
var PASSWORD = 'BP1234';
var DASHBOARD_SHEET = '대시보드';

// =========================================================
//  1. 웹 앱 진입점 — 대시보드 HTML 페이지 서빙
// =========================================================

function doGet(e) {
  var params = (e && e.parameter) || {};

  // API 모드: action 파라미터가 있으면 JSON 응답
  if (params.action) {
    var data = {};
    if (params.data) { try { data = JSON.parse(params.data); } catch(err) {} }
    var result = processApiRequest(params.action, params.token, data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 페이지 모드: 대시보드 HTML 서빙
  return HtmlService.createHtmlOutput(getDashboardHtml())
    .setTitle('경영전략 파트너 연락처')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
}

function processApiRequest(action, token, data) {
  if (action === 'ping') return { success: true };
  if (!isValidHash(token)) return { success: false, error: 'UNAUTHORIZED' };
  switch(action) {
    case 'validate': return { success: true };
    case 'getContacts': return { success: true, contacts: getAllContacts() };
    case 'getCategories': return { success: true, categories: getCategoryList() };
    case 'addContact': return addContact(data);
    default: return { success: false, error: 'UNKNOWN' };
  }
}

// =========================================================
//  2. 시트 내 커스텀 메뉴
// =========================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('연락처 관리')
    .addItem('신규 연락처 등록', 'showAddContactDialog')
    .addItem('대시보드로 이동', 'goToDashboard')
    .addToUi();
}

function goToDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET);
  if (sheet) ss.setActiveSheet(sheet);
}

function showAddContactDialog() {
  var html = HtmlService.createHtmlOutput(getDialogHtml())
    .setWidth(480).setHeight(580).setTitle('신규 연락처 등록');
  SpreadsheetApp.getUi().showModalDialog(html, '신규 연락처 등록');
}

// =========================================================
//  3. 보안 API — 웹 대시보드에서 google.script.run 으로 호출
// =========================================================

/** 암호 해시 검증 */
function validatePassword(hash) {
  return { success: isValidHash(hash) };
}

/** 전체 연락처 조회 (인증 필수) */
function getContactsSecure(hash) {
  if (!isValidHash(hash)) return { success: false, error: 'AUTH' };
  return { success: true, contacts: getAllContacts() };
}

/** 구분 목록 조회 (인증 필수) */
function getCategoriesSecure(hash) {
  if (!isValidHash(hash)) return { success: false, error: 'AUTH' };
  return { success: true, categories: getCategoryList() };
}

/** 연락처 추가 (인증 필수) */
function addContactSecure(hash, data) {
  if (!isValidHash(hash)) return { success: false, error: 'AUTH' };
  return addContact(data);
}

// =========================================================
//  4. 인증
// =========================================================

function isValidHash(hash) {
  if (!hash) return false;
  return hash === computeHash(PASSWORD);
}

function computeHash(text) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8
  );
  return raw.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

// =========================================================
//  5. 비즈니스 로직
// =========================================================

function getCategoryList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var categories = [];
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getName() !== DASHBOARD_SHEET) categories.push(sheet.getName());
  });
  return categories;
}

function getAllContacts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allContacts = [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name === DASHBOARD_SHEET) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return;
    var data = sheet.getRange(4, 3, lastRow - 3, 7).getValues();
    data.forEach(function(row) {
      if (row[0]) {
        allContacts.push({
          category: name, company: String(row[0]), name: String(row[1]),
          title: String(row[2]), team: String(row[3]), phone: String(row[4]),
          email: String(row[5]), note: String(row[6] || '')
        });
      }
    });
  });
  return allContacts;
}

function addContact(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var categoryName = data.category;

  if (data.isNewCategory && data.newCategoryName) {
    categoryName = String(data.newCategoryName).trim();
    if (ss.getSheetByName(categoryName)) {
      return { success: false, message: '이미 존재하는 구분입니다: ' + categoryName };
    }
    var newSheet = ss.insertSheet(categoryName);
    newSheet.getRange('B2:I2').setValues([['No','기업','성함','직책','조직/팀','연락처','','비고']]);
    newSheet.getRange('G3:H3').setValues([['M','E']]);
    newSheet.getRange('B2:I2').setFontWeight('bold');
    updateDashboardQuery(ss);
  }

  var sheet = ss.getSheetByName(categoryName);
  if (!sheet) return { success: false, message: '시트를 찾을 수 없습니다: ' + categoryName };

  var lastRow = sheet.getLastRow();
  var nextRow = (lastRow < 4) ? 4 : lastRow + 1;
  sheet.getRange(nextRow, 3, 1, 7).setValues([[
    String(data.company||''), String(data.name||''), String(data.title||''),
    String(data.team||''), String(data.phone||''), String(data.email||''),
    String(data.note||'')
  ]]);

  return { success: true, message: categoryName + '에 ' + data.name + ' 연락처가 등록되었습니다!' };
}

function updateDashboardQuery(ss) {
  var categories = [];
  ss.getSheets().forEach(function(s) {
    if (s.getName() !== DASHBOARD_SHEET) categories.push(s.getName());
  });
  var parts = categories.map(function(c) {
    var sn = (c.indexOf('/')>=0||c.indexOf(' ')>=0) ? "'"+c+"'" : c;
    return 'ARRAYFORMULA(IF('+sn+'!C4:C1000<>"","'+c+'","")),'+sn+'!C4:I1000';
  });
  var f = '=QUERY({'+parts.join(';')+'},"SELECT * WHERE Col1<>\'\' ORDER BY Col1, Col2",0)';
  var d = ss.getSheetByName(DASHBOARD_SHEET);
  if (d) d.getRange('A9').setFormula(f);
}

// =========================================================
//  6. 시트 내 등록 다이얼로그 HTML (스프레드시트에서 직접 사용)
// =========================================================

function getDialogHtml() {
  var cats = getCategoryList();
  var opts = cats.map(function(c){return '<option value="'+c+'">'+c+'</option>';}).join('');
  return '<!DOCTYPE html><html><head><style>'
    +'*{box-sizing:border-box;margin:0;padding:0}'
    +'body{font-family:"Google Sans",Arial,sans-serif;padding:24px}'
    +'h2{color:#1a73e8;margin-bottom:20px;font-size:18px}'
    +'.fg{margin-bottom:14px}label{display:block;font-size:13px;font-weight:500;color:#5f6368;margin-bottom:4px}'
    +'input,select{width:100%;padding:10px 12px;border:1px solid #dadce0;border-radius:8px;font-size:14px;outline:none}'
    +'input:focus,select:focus{border-color:#1a73e8}'
    +'.row{display:flex;gap:12px}.row .fg{flex:1}'
    +'.cb{display:flex;align-items:center;gap:8px;margin-top:8px}.cb input{width:auto}.cb label{margin:0;color:#1a73e8;cursor:pointer}'
    +'.nc{display:none;margin-top:8px}.nc.show{display:block}'
    +'.btns{display:flex;gap:10px;margin-top:20px}'
    +'.btn{flex:1;padding:12px;border:none;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer}'
    +'.btn-p{background:#1a73e8;color:#fff}.btn-s{background:#f1f3f4;color:#5f6368}'
    +'.msg{padding:10px;border-radius:8px;margin-top:14px;font-size:13px;display:none}'
    +'.msg.ok{display:block;background:#e6f4ea;color:#137333}.msg.err{display:block;background:#fce8e6;color:#c5221f}'
    +'</style></head><body>'
    +'<h2>신규 연락처 등록</h2>'
    +'<div class="fg"><label>구분 *</label><select id="cat">'+opts+'</select>'
    +'<div class="cb"><input type="checkbox" id="nc" onchange="tg()"><label for="nc">새 구분 추가</label></div>'
    +'<div class="nc" id="ncg"><input type="text" id="ncn" placeholder="새 구분 이름"></div></div>'
    +'<div class="fg"><label>기업 *</label><input id="co" placeholder="기업명"></div>'
    +'<div class="row"><div class="fg"><label>성함 *</label><input id="nm" placeholder="이름"></div>'
    +'<div class="fg"><label>직책</label><input id="ti" placeholder="직책"></div></div>'
    +'<div class="fg"><label>조직/팀</label><input id="tm" placeholder="조직 또는 팀명"></div>'
    +'<div class="row"><div class="fg"><label>연락처 (M)</label><input id="ph" placeholder="010-0000-0000"></div>'
    +'<div class="fg"><label>연락처 (E)</label><input id="em" placeholder="email@company.com"></div></div>'
    +'<div class="fg"><label>비고</label><input id="no" placeholder="메모"></div>'
    +'<div class="btns"><button class="btn btn-s" onclick="google.script.host.close()">취소</button>'
    +'<button class="btn btn-p" id="sb" onclick="sub()">등록</button></div>'
    +'<div class="msg" id="mg"></div>'
    +'<script>'
    +'function tg(){var c=document.getElementById("nc").checked;document.getElementById("ncg").className=c?"nc show":"nc";document.getElementById("cat").disabled=c;}'
    +'function sub(){'
    +'var co=document.getElementById("co").value.trim(),nm=document.getElementById("nm").value.trim();'
    +'if(!co||!nm){sm("기업명과 성함은 필수입니다","err");return;}'
    +'var n=document.getElementById("nc").checked;'
    +'if(n&&!document.getElementById("ncn").value.trim()){sm("새 구분 이름을 입력하세요","err");return;}'
    +'document.getElementById("sb").disabled=true;document.getElementById("sb").textContent="등록 중...";'
    +'google.script.run.withSuccessHandler(function(r){'
    +'if(r.success){sm(r.message,"ok");setTimeout(function(){google.script.host.close();},1200);}'
    +'else{sm(r.message,"err");document.getElementById("sb").disabled=false;document.getElementById("sb").textContent="등록";}}'
    +').withFailureHandler(function(e){'
    +'sm("오류: "+e.message,"err");document.getElementById("sb").disabled=false;document.getElementById("sb").textContent="등록";})'
    +'.addContact({'
    +'category:document.getElementById("cat").value,isNewCategory:n,newCategoryName:document.getElementById("ncn").value,'
    +'company:co,name:nm,title:document.getElementById("ti").value.trim(),'
    +'team:document.getElementById("tm").value.trim(),phone:document.getElementById("ph").value.trim(),'
    +'email:document.getElementById("em").value.trim(),note:document.getElementById("no").value.trim()});}'
    +'function sm(t,c){var e=document.getElementById("mg");e.textContent=t;e.className="msg "+c;}'
    +'<\/script></body></html>';
}

// =========================================================
//  7. 웹 대시보드 HTML (전체 페이지)
// =========================================================

function getDashboardHtml() {
  return '\
<!DOCTYPE html>\
<html lang="ko">\
<head>\
<meta charset="UTF-8">\
<style>\
:root{--p:#2563eb;--ph:#1d4ed8;--pl:#dbeafe;--s:#059669;--sl:#d1fae5;--d:#dc2626;--dl:#fee2e2;\
--g50:#f9fafb;--g100:#f3f4f6;--g200:#e5e7eb;--g300:#d1d5db;--g400:#9ca3af;--g500:#6b7280;\
--g600:#4b5563;--g700:#374151;--g800:#1f2937;--r:12px;--rs:8px}\
*{box-sizing:border-box;margin:0;padding:0}\
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:var(--g50);color:var(--g800);min-height:100vh}\
.scr{display:none}.scr.on{display:flex}\
#LS{justify-content:center;align-items:center;min-height:100vh;background:linear-gradient(135deg,#1e3a5f,#2563eb)}\
.lc{background:#fff;border-radius:20px;padding:48px 40px;width:400px;max-width:90vw;text-align:center;box-shadow:0 25px 50px rgba(0,0,0,.2)}\
.li{width:64px;height:64px;background:var(--pl);border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 24px}\
.li svg{width:28px;height:28px;stroke:var(--p);fill:none;stroke-width:2}\
.lc h1{font-size:22px;margin-bottom:8px}.lc p{font-size:14px;color:var(--g500);margin-bottom:24px}\
.lc input{width:100%;padding:14px 16px;border:2px solid var(--g200);border-radius:var(--rs);font-size:16px;outline:none;text-align:center;letter-spacing:4px;transition:border .2s}\
.lc input:focus{border-color:var(--p)}\
.bl{width:100%;padding:14px;background:var(--p);color:#fff;border:none;border-radius:var(--rs);font-size:16px;font-weight:600;cursor:pointer;margin-top:16px;transition:background .2s}\
.bl:hover{background:var(--ph)}.bl:disabled{opacity:.5;cursor:not-allowed}\
.le{color:var(--d);font-size:13px;margin-top:12px;min-height:20px}\
#DS{flex-direction:column;min-height:100vh}\
.tb{background:#fff;border-bottom:1px solid var(--g200);padding:16px 32px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}\
.tb h1{font-size:18px}\
.ta{display:flex;gap:12px;align-items:center}\
.btn{padding:10px 20px;border-radius:var(--rs);font-size:14px;font-weight:500;cursor:pointer;border:none;transition:all .2s;display:inline-flex;align-items:center;gap:6px}\
.bp{background:var(--p);color:#fff}.bp:hover{background:var(--ph)}\
.bo{background:#fff;color:var(--g600);border:1px solid var(--g300)}.bo:hover{background:var(--g50)}\
.bs{padding:6px 12px;font-size:13px}\
.mc{padding:24px 32px;flex:1;max-width:1200px;width:100%;margin:0 auto}\
.sg{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px}\
.sc{background:#fff;border-radius:var(--r);padding:16px 20px;border:1px solid var(--g200);cursor:pointer;transition:all .2s}\
.sc:hover{border-color:var(--p);box-shadow:0 2px 8px rgba(37,99,235,.1)}\
.sc.ac{border-color:var(--p);background:var(--pl)}\
.sc .sl{font-size:12px;color:var(--g500);font-weight:500}.sc .sv{font-size:28px;font-weight:700;color:var(--g800);margin-top:4px}\
.sc.ac .sv{color:var(--p)}.sc.tt{border-color:var(--p)}.sc.tt .sv{color:var(--p)}\
.sb{margin-bottom:20px;position:relative}\
.sb input{width:100%;padding:12px 16px 12px 44px;border:2px solid var(--g200);border-radius:var(--rs);font-size:14px;outline:none;background:#fff;transition:border .2s}\
.sb input:focus{border-color:var(--p)}\
.sb svg{position:absolute;left:14px;top:50%;transform:translateY(-50%);width:18px;height:18px;stroke:var(--g400);fill:none;stroke-width:2}\
.tw{background:#fff;border-radius:var(--r);border:1px solid var(--g200);overflow:hidden}\
table{width:100%;border-collapse:collapse}\
thead th{padding:14px 16px;text-align:left;font-size:12px;font-weight:600;color:var(--g500);text-transform:uppercase;background:var(--g50);border-bottom:2px solid var(--g200);white-space:nowrap}\
tbody td{padding:14px 16px;font-size:14px;border-bottom:1px solid var(--g100)}\
tbody tr:hover{background:var(--g50)}\
.cb2{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:500}\
.b-ir{background:#dbeafe;color:#1d4ed8}.b-cd{background:#fef3c7;color:#92400e}\
.b-inv{background:#d1fae5;color:#065f46}.b-mob{background:#ede9fe;color:#5b21b6}\
.b-pg{background:#fce7f3;color:#9d174d}.b-ins{background:#fed7aa;color:#9a3412}.b-def{background:var(--g100);color:var(--g600)}\
.cp,.ce{font-family:"SF Mono","Fira Code",monospace;font-size:13px;color:var(--g600)}\
.es{text-align:center;padding:60px 20px;color:var(--g400)}.es svg{width:48px;height:48px;stroke:var(--g300);fill:none;margin-bottom:16px}.es p{font-size:15px}\
.mo{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:100;justify-content:center;align-items:center;backdrop-filter:blur(4px)}\
.mo.on{display:flex}\
.md{background:#fff;border-radius:16px;width:520px;max-width:95vw;max-height:90vh;overflow-y:auto;box-shadow:0 25px 50px rgba(0,0,0,.2)}\
.mh{padding:24px 28px 0;display:flex;justify-content:space-between;align-items:center}.mh h2{font-size:18px}\
.mx{width:32px;height:32px;border-radius:8px;border:none;background:var(--g100);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:18px;color:var(--g500)}.mx:hover{background:var(--g200)}\
.mb{padding:24px 28px 28px}\
.fg{margin-bottom:16px}.fg label{display:block;font-size:13px;font-weight:500;color:var(--g600);margin-bottom:5px}\
.fg label .rq{color:var(--d)}\
.fg input,.fg select{width:100%;padding:10px 14px;border:1.5px solid var(--g200);border-radius:var(--rs);font-size:14px;outline:none;transition:border .2s}\
.fg input:focus,.fg select:focus{border-color:var(--p)}\
.fr{display:flex;gap:12px}.fr .fg{flex:1}\
.ck{display:flex;align-items:center;gap:8px;margin-top:8px}.ck input[type=checkbox]{width:16px;height:16px;accent-color:var(--p)}\
.ck label{font-size:13px;color:var(--p);cursor:pointer;margin:0}\
.ni{margin-top:8px;display:none}.ni.sh{display:block}\
.ma{display:flex;gap:10px;margin-top:24px}.ma .btn{flex:1;padding:12px;justify-content:center}\
.toast{position:fixed;bottom:24px;right:24px;padding:14px 24px;border-radius:var(--rs);font-size:14px;font-weight:500;z-index:200;transform:translateY(100px);opacity:0;transition:all .3s;box-shadow:0 8px 24px rgba(0,0,0,.15)}\
.toast.sh{transform:translateY(0);opacity:1}.toast.ok{background:var(--s);color:#fff}.toast.er{background:var(--d);color:#fff}\
.lo{display:none;position:fixed;inset:0;background:rgba(255,255,255,.8);z-index:150;justify-content:center;align-items:center}.lo.on{display:flex}\
.sp{width:40px;height:40px;border:4px solid var(--g200);border-top-color:var(--p);border-radius:50%;animation:spin .8s linear infinite}\
@keyframes spin{to{transform:rotate(360deg)}}\
@media(max-width:768px){.tb{padding:12px 16px;flex-wrap:wrap;gap:12px}.tb h1{font-size:16px}.mc{padding:16px}.sg{grid-template-columns:repeat(3,1fr);gap:8px}.sc{padding:12px}.sc .sv{font-size:22px}.tw{overflow-x:auto}table{min-width:700px}.fr{flex-direction:column;gap:0}}\
@media(max-width:480px){.sg{grid-template-columns:repeat(2,1fr)}}\
</style>\
</head>\
<body>\
<div id="LS" class="scr on">\
<div class="lc">\
<div class="li"><svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg></div>\
<h1>Partner Contacts</h1>\
<p>접속 암호를 입력하세요</p>\
<input type="password" id="pwi" placeholder="Password" autocomplete="off">\
<button class="bl" id="lbtn" onclick="doLogin()">접속</button>\
<div class="le" id="lerr"></div>\
</div>\
</div>\
<div id="DS" class="scr">\
<div class="tb">\
<h1>Partner Contacts</h1>\
<div class="ta">\
<button class="btn bp" onclick="showAdd()"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>신규 연락처 등록</button>\
<button class="btn bo bs" onclick="reload()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>새로고침</button>\
<button class="btn bo bs" onclick="doLogout()">로그아웃</button>\
</div>\
</div>\
<div class="mc">\
<div class="sg" id="sGrid"></div>\
<div class="sb"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>\
<input type="text" id="sInp" placeholder="이름, 기업, 연락처로 검색..." oninput="renderC()"></div>\
<div class="tw"><table><thead><tr><th>구분</th><th>기업</th><th>성함</th><th>직책</th><th>조직/팀</th><th>연락처(M)</th><th>연락처(E)</th><th>비고</th></tr></thead>\
<tbody id="cTb"></tbody></table></div>\
</div>\
</div>\
<div class="mo" id="amo" onclick="if(event.target===this)closeAdd()">\
<div class="md">\
<div class="mh"><h2>신규 연락처 등록</h2><button class="mx" onclick="closeAdd()">&times;</button></div>\
<div class="mb">\
<div class="fg"><label>구분 <span class="rq">*</span></label><select id="mCat"></select>\
<div class="ck"><input type="checkbox" id="mNc" onchange="tgNc()"><label for="mNc">새 구분 추가</label></div>\
<div class="ni" id="mNcI"><input type="text" id="mNcN" placeholder="새 구분 이름 (예: Legal)"></div></div>\
<div class="fg"><label>기업 <span class="rq">*</span></label><input type="text" id="mCo" placeholder="기업명"></div>\
<div class="fr"><div class="fg"><label>성함 <span class="rq">*</span></label><input type="text" id="mNm" placeholder="이름"></div>\
<div class="fg"><label>직책</label><input type="text" id="mTi" placeholder="직책"></div></div>\
<div class="fg"><label>조직/팀</label><input type="text" id="mTm" placeholder="조직 또는 팀명"></div>\
<div class="fr"><div class="fg"><label>연락처 (M)</label><input type="tel" id="mPh" placeholder="010-0000-0000"></div>\
<div class="fg"><label>연락처 (E)</label><input type="email" id="mEm" placeholder="email@company.com"></div></div>\
<div class="fg"><label>비고</label><input type="text" id="mNo" placeholder="메모"></div>\
<div class="ma"><button class="btn bo" onclick="closeAdd()">취소</button><button class="btn bp" id="subBtn" onclick="doSubmit()">등록</button></div>\
</div></div></div>\
<div class="lo" id="ldOv"><div class="sp"></div></div>\
<div class="toast" id="tst"></div>\
<script>\
var contacts=[],cats=[],token="",curF="전체";\
function sha256(m){var e=new TextEncoder().encode(m);return crypto.subtle.digest("SHA-256",e).then(function(b){return Array.from(new Uint8Array(b)).map(function(x){return x.toString(16).padStart(2,"0")}).join("")})}\
function srv(fn){var args=Array.prototype.slice.call(arguments,1);return new Promise(function(ok,fail){var r=google.script.run.withSuccessHandler(ok).withFailureHandler(fail);r[fn].apply(r,args)})}\
function doLogin(){var pw=document.getElementById("pwi").value,err=document.getElementById("lerr"),btn=document.getElementById("lbtn");\
if(!pw){err.textContent="암호를 입력하세요";return}btn.disabled=true;btn.textContent="확인 중...";err.textContent="";\
sha256(pw).then(function(h){token=h;return srv("validatePassword",h)}).then(function(r){\
if(r.success){sessionStorage.setItem("_t",token);sw("DS");loadAll()}\
else{err.textContent="암호가 올바르지 않습니다";token=""}\
btn.disabled=false;btn.textContent="접속";\
}).catch(function(){err.textContent="접속 실패. 다시 시도하세요.";token="";btn.disabled=false;btn.textContent="접속"})}\
function doLogout(){token="";sessionStorage.removeItem("_t");contacts=[];cats=[];sw("LS");document.getElementById("pwi").value=""}\
function tryAuto(){var s=sessionStorage.getItem("_t");if(!s)return;token=s;\
srv("validatePassword",s).then(function(r){if(r.success){sw("DS");loadAll()}else sessionStorage.removeItem("_t")}).catch(function(){sessionStorage.removeItem("_t")})}\
function loadAll(){ld(true);Promise.all([srv("getContactsSecure",token),srv("getCategoriesSecure",token)]).then(function(r){\
if(r[0].success)contacts=r[0].contacts;if(r[1].success)cats=r[1].categories;renderS();renderC();ld(false);\
}).catch(function(){toast("데이터 로드 실패","er");ld(false)})}\
function reload(){loadAll().then&&loadAll();toast("새로고침 완료","ok")}\
function renderS(){var g=document.getElementById("sGrid"),cn={},t=contacts.length;cats.forEach(function(c){cn[c]=0});\
contacts.forEach(function(c){cn[c.category]=(cn[c.category]||0)+1});\
var h="<div class=\\"sc tt"+(curF==="전체"?" ac":"")+ "\\" onclick=\\"setF(\'전체\')\\"><div class=\\"sl\\">전체</div><div class=\\"sv\\">"+t+"</div></div>";\
cats.forEach(function(c){h+="<div class=\\"sc"+(curF===c?" ac":"")+ "\\" onclick=\\"setF(\'"+esc(c)+"\')\\"><div class=\\"sl\\">"+esc(c)+"</div><div class=\\"sv\\">"+(cn[c]||0)+"</div></div>"});\
g.innerHTML=h}\
function renderC(){var tb=document.getElementById("cTb"),q=(document.getElementById("sInp").value||"").toLowerCase().trim(),f=contacts;\
if(curF!=="전체")f=f.filter(function(c){return c.category===curF});\
if(q)f=f.filter(function(c){return(c.company||"").toLowerCase().indexOf(q)>=0||(c.name||"").toLowerCase().indexOf(q)>=0||(c.phone||"").indexOf(q)>=0||(c.email||"").toLowerCase().indexOf(q)>=0||(c.team||"").toLowerCase().indexOf(q)>=0||(c.title||"").toLowerCase().indexOf(q)>=0});\
if(!f.length){tb.innerHTML="<tr><td colspan=\\"8\\"><div class=\\"es\\"><svg viewBox=\\"0 0 24 24\\" stroke-width=\\"1.5\\"><path d=\\"M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2\\"/><circle cx=\\"9\\" cy=\\"7\\" r=\\"4\\"/></svg><p>"+(q?"검색 결과가 없습니다":"등록된 연락처가 없습니다")+"</p></div></td></tr>";return}\
var h="";f.forEach(function(c){var bc=badge(c.category);h+="<tr><td><span class=\\"cb2 "+bc+"\\">"+esc(c.category)+"</span></td><td><strong>"+esc(c.company)+"</strong></td><td>"+esc(c.name)+"</td><td>"+esc(c.title)+"</td><td>"+esc(c.team)+"</td><td class=\\"cp\\">"+esc(c.phone)+"</td><td class=\\"ce\\">"+esc(c.email)+"</td><td>"+esc(c.note)+"</td></tr>"});\
tb.innerHTML=h}\
function setF(c){curF=c;renderS();renderC()}\
function badge(c){var m={"IR":"b-ir","CorpDev":"b-cd","Invest":"b-inv","Mobility":"b-mob","PG/Payment":"b-pg","Insurance":"b-ins"};return m[c]||"b-def"}\
function showAdd(){var s=document.getElementById("mCat");s.innerHTML=cats.map(function(c){return"<option value=\\""+esc(c)+"\\">"+esc(c)+"</option>"}).join("");s.disabled=false;\
document.getElementById("mNc").checked=false;document.getElementById("mNcI").classList.remove("sh");\
["mCo","mNm","mTi","mTm","mPh","mEm","mNo","mNcN"].forEach(function(id){document.getElementById(id).value=""});\
document.getElementById("subBtn").disabled=false;document.getElementById("subBtn").textContent="등록";\
document.getElementById("amo").classList.add("on")}\
function closeAdd(){document.getElementById("amo").classList.remove("on")}\
function tgNc(){var ch=document.getElementById("mNc").checked;document.getElementById("mNcI").classList.toggle("sh",ch);document.getElementById("mCat").disabled=ch}\
function doSubmit(){var nw=document.getElementById("mNc").checked,co=document.getElementById("mCo").value.trim(),nm=document.getElementById("mNm").value.trim();\
if(!co||!nm){toast("기업명과 성함은 필수입니다","er");return}\
if(nw&&!document.getElementById("mNcN").value.trim()){toast("새 구분 이름을 입력하세요","er");return}\
var btn=document.getElementById("subBtn");btn.disabled=true;btn.textContent="등록 중...";\
var d={category:document.getElementById("mCat").value,isNewCategory:nw,newCategoryName:document.getElementById("mNcN").value.trim(),\
company:co,name:nm,title:document.getElementById("mTi").value.trim(),team:document.getElementById("mTm").value.trim(),\
phone:document.getElementById("mPh").value.trim(),email:document.getElementById("mEm").value.trim(),note:document.getElementById("mNo").value.trim()};\
srv("addContactSecure",token,d).then(function(r){if(r.success){toast(r.message,"ok");closeAdd();loadAll()}else{toast(r.message||"등록 실패","er");btn.disabled=false;btn.textContent="등록"}\
}).catch(function(){toast("오류가 발생했습니다","er");btn.disabled=false;btn.textContent="등록"})}\
function sw(id){document.querySelectorAll(".scr").forEach(function(s){s.classList.remove("on")});document.getElementById(id).classList.add("on")}\
function ld(s){document.getElementById("ldOv").classList.toggle("on",s)}\
function toast(m,t){var e=document.getElementById("tst");e.textContent=m;e.className="toast "+t+" sh";setTimeout(function(){e.classList.remove("sh")},3000)}\
function esc(s){if(!s)return"";var d=document.createElement("div");d.textContent=s;return d.innerHTML}\
document.getElementById("pwi").addEventListener("keydown",function(e){if(e.key==="Enter")doLogin()});\
tryAuto();\
<\/script>\
</body>\
</html>';
}
