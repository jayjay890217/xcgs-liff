/**
 * 旭馳車業｜LIFF 會員綁定後端
 * LIFF → Apps Script（驗證 ID Token）→ Google 試算表
 *
 * 部署：擴充功能 → Apps Script → 貼上本檔 → 部署 → 新增部署作業
 *       類型「網頁應用程式」，執行身分「我」，存取權「所有人」
 */

// ============================================================
// 設定
// ============================================================
const LIFF_CHANNEL_ID = '2011218129';   // LIFF ID 的「-」前面那段 = LINE Login Channel ID
const SHEET_NAME = '會員綁定';

const HEADERS = [
  'LINE UserID', 'LINE 名稱', '姓名', '手機', '車牌', '車種', '關係',
  '首次綁定時間', '最後更新時間', '客戶ID', '比對結果'
];
const COL = { userId: 1, displayName: 2, name: 3, phone: 4, plate: 5, model: 6, relationship: 7, createdAt: 8, updatedAt: 9, customerId: 10, match: 11 };

// 旭馳APP資料庫（AppSheet 旭馳雲端系統）的「客戶」分頁：一列＝一台車
const APP_DB_ID = '1zjWNnVYUlILxes3Gn-PTRJpO1kgOmjeogGX9hK-4UEM';
const APP_SHEET = '客戶';
const RELATIONSHIPS = ['車主本人', '家人', '朋友', '其他'];

// 綁定成功後換成「已登入」圖文選單
// Messaging API 的 Channel access token 放在「專案設定 → 指令碼屬性」，名稱 LINE_CHANNEL_ACCESS_TOKEN（不要寫在程式碼裡）
const MEMBER_MENU_ALIAS = 'xuchi-member';

// ============================================================
// 入口
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const user = verifyIdToken_(body.idToken);

    switch (body.action) {
      case 'get':
        return json_({ ok: true, member: findMember_(user.userId) });
      case 'bind':
        return json_(bind_(user, body));
      default:
        return json_({ ok: false, error: '未知的操作' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err.userMessage || '系統忙碌，請稍後再試', code: err.code || 'ERROR' });
  }
}

// 讓你直接用瀏覽器打開網址確認有部署成功
function doGet() {
  return json_({ ok: true, service: 'xcgs-liff-bind' });
}

// ============================================================
// 向 LINE 驗證 ID Token，取得「真的」userId
// ============================================================
function verifyIdToken_(idToken) {
  if (!idToken) throw userError_('缺少登入資訊，請從 LINE 重新開啟', 'NO_TOKEN');

  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: LIFF_CHANNEL_ID },
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !data.sub) {
    const expired = /expired/i.test(data.error_description || '');
    throw userError_(
      expired ? '登入已過期，請重新開啟' : 'LINE 身分驗證失敗，請從 LINE 重新開啟',
      expired ? 'TOKEN_EXPIRED' : 'TOKEN_INVALID'
    );
  }
  return { userId: data.sub, displayName: data.name || '' };
}

// ============================================================
// 綁定（同一個 LINE 帳號只會有一列，重複送出＝更新）
// ============================================================
function bind_(user, body) {
  const m = {
    name: clean_(body.name, 30),
    phone: String(body.phone || '').replace(/\D/g, '').replace(/^8869/, '09'),
    plate: String(body.plate || '').toUpperCase().replace(/[^A-Z0-9-]/g, ''),
    model: clean_(body.model, 40),
    relationship: RELATIONSHIPS.indexOf(body.relationship) >= 0 ? body.relationship : '其他'
  };

  if (!m.name) throw userError_('請填寫姓名');
  if (!/^09\d{8}$/.test(m.phone)) throw userError_('手機號碼格式不正確');
  if (!/^[A-Z0-9]{2,4}-[A-Z0-9]{2,4}$/.test(m.plate)) throw userError_('車牌格式不正確');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, user.userId);
    const now = new Date();

    // 對應旭馳雲端車輛（失敗不影響綁定本身）
    let link = { customerId: '', match: '' };
    try { link = linkToApp_(user.userId, m); } catch (err) { console.error(err); link.match = '比對失敗：' + err.message; }

    const values = [
      user.userId, user.displayName, m.name, m.phone, m.plate, m.model, m.relationship,
      row ? sheet.getRange(row, COL.createdAt).getValue() : now,
      now, link.customerId, link.match
    ];

    const target = row || sheet.getLastRow() + 1;
    sheet.getRange(target, 1, 1, HEADERS.length).setValues([values]);

    // 換成會員選單（失敗不影響綁定本身）
    let menu = false;
    try { menu = linkMemberMenu_(user.userId); } catch (err) { console.error(err); }

    return { ok: true, updated: !!row, member: m, menu: menu };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 對應旭馳雲端：用車牌找車，車主本人且手機相符才寫入 LINE_User_ID
// ============================================================
function normPlate_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function normPhone_(s) { return String(s || '').replace(/\D/g, ''); }

function linkToApp_(userId, m) {
  const sh = SpreadsheetApp.openById(APP_DB_ID).getSheetByName(APP_SHEET);
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0];
  const c = { id: h.indexOf('客戶ID'), phone: h.indexOf('手機'), plate: h.indexOf('車牌'), line: h.indexOf('LINE_User_ID') };
  if (c.id < 0 || c.plate < 0 || c.line < 0) throw new Error('旭馳雲端「客戶」分頁欄位名稱不符');

  const want = normPlate_(m.plate);
  const phone = normPhone_(m.phone);
  let hit = -1;
  for (let i = 1; i < v.length; i++) if (normPlate_(v[i][c.plate]) === want) { hit = i; break; }

  // 這個 LINE 帳號之前綁在「別支手機」的車（真的換人或換號碼）才清掉；同一支手機的多台車全部保留
  for (let i = 1; i < v.length; i++) {
    if (v[i][c.line] === userId && (c.phone < 0 || normPhone_(v[i][c.phone]) !== phone)) {
      sh.getRange(i + 1, c.line + 1).setValue('');
    }
  }

  if (hit < 0) return { customerId: '', match: '找不到車牌（新車待建檔）' };

  const r = v[hit], customerId = r[c.id];
  if (m.relationship !== '車主本人') return { customerId: customerId, match: '已對應車輛（' + m.relationship + '，未寫入車主資料）' };

  const appPhone = normPhone_(r[c.phone]);
  if (appPhone && appPhone !== phone) return { customerId: customerId, match: '待確認：手機與旭馳雲端不同' };
  if (r[c.line] && r[c.line] !== userId) return { customerId: customerId, match: '待確認：這台車已綁定其他 LINE 帳號' };

  // 車牌＋手機驗證過 → 這台車寫入 LINE
  sh.getRange(hit + 1, c.line + 1).setValue(userId);
  if (!appPhone && c.phone >= 0) sh.getRange(hit + 1, c.phone + 1).setNumberFormat('@').setValue(phone);

  // 同一支手機名下的其他車，LINE 欄空白的一起寫入（已綁別的 LINE 的不動）
  let more = 0;
  if (c.phone >= 0) {
    for (let i = 1; i < v.length; i++) {
      if (i === hit || normPhone_(v[i][c.phone]) !== phone) continue;
      if (v[i][c.line] && v[i][c.line] !== userId) continue;
      if (v[i][c.line] !== userId) sh.getRange(i + 1, c.line + 1).setValue(userId);
      more++;
    }
  }
  const tail = more ? '（同手機共 ' + (more + 1) + ' 台）' : '';
  return { customerId: customerId, match: (appPhone ? '已對應並寫入 LINE' : '已對應並寫入 LINE（補上手機）') + tail };
}

// ============================================================
// 每天自動補：日後新建檔的車，手機跟已驗證的會員相同 → 自動寫入 LINE
// （先在編輯器執行一次 setupDailySync 建立排程）
// ============================================================
function syncByPhone() {
  const members = getSheet_();
  const last = members.getLastRow();
  if (last < 2) return 0;
  const phoneToUser = {};
  members.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues().forEach(function (r) {
    // 只信任「車主本人」且已經用車牌＋手機驗證成功的
    if (r[COL.relationship - 1] === '車主本人' && /^已對應並寫入 LINE/.test(r[COL.match - 1])) {
      phoneToUser[normPhone_(r[COL.phone - 1])] = r[COL.userId - 1];
    }
  });

  const sh = SpreadsheetApp.openById(APP_DB_ID).getSheetByName(APP_SHEET);
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0];
  const c = { phone: h.indexOf('手機'), plate: h.indexOf('車牌'), line: h.indexOf('LINE_User_ID') };
  if (c.phone < 0 || c.line < 0) throw new Error('旭馳雲端「客戶」分頁欄位名稱不符');

  let n = 0;
  for (let i = 1; i < v.length; i++) {
    if (v[i][c.line]) continue;
    const uid = phoneToUser[normPhone_(v[i][c.phone])];
    if (!uid) continue;
    sh.getRange(i + 1, c.line + 1).setValue(uid);
    n++;
  }
  Logger.log('同手機自動補上 LINE：' + n + ' 台');
  return n;
}

function setupDailySync() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncByPhone') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncByPhone').timeBased().everyDays(1).atHour(3).inTimezone('Asia/Taipei').create();
  syncByPhone();
  Logger.log('已建立每日凌晨 3 點自動同步');
}

// 已經綁定過的會員補跑一次比對（在編輯器手動執行）
function relinkAll() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  rows.forEach(function (v, i) {
    const m = { phone: v[COL.phone - 1], plate: v[COL.plate - 1], relationship: v[COL.relationship - 1] };
    let link;
    try { link = linkToApp_(v[COL.userId - 1], m); } catch (err) { link = { customerId: '', match: '比對失敗：' + err.message }; }
    sheet.getRange(i + 2, COL.customerId, 1, 2).setValues([[link.customerId, link.match]]);
  });
  Logger.log('完成 ' + rows.length + ' 筆');
  return rows.length;
}

// ============================================================
// 圖文選單：把「已登入」選單套用到這位顧客
// ============================================================
function lineToken_() {
  return PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '';
}

// 用別名找出目前的會員選單 ID（選單重建時 member_menu.py 會更新別名，這裡不用改）
function memberMenuId_(token) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('memberMenuId');
  if (hit) return hit;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/alias/' + MEMBER_MENU_ALIAS, {
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('找不到會員選單別名 ' + MEMBER_MENU_ALIAS + '（HTTP ' + res.getResponseCode() + '）');
  const id = JSON.parse(res.getContentText()).richMenuId;
  cache.put('memberMenuId', id, 3600);
  return id;
}

function linkMemberMenu_(userId) {
  const token = lineToken_();
  if (!token) { console.warn('未設定 LINE_CHANNEL_ACCESS_TOKEN，略過換選單'); return false; }
  const menuId = memberMenuId_(token);
  const res = UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/user/' + encodeURIComponent(userId) + '/richmenu/' + menuId,
    { method: 'post', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) {
    CacheService.getScriptCache().remove('memberMenuId');  // 下次重新查別名
    console.error('換選單失敗 ' + res.getResponseCode() + ' ' + res.getContentText());
    return false;
  }
  return true;
}

// 已經綁定過的會員，一次全部換成會員選單（在編輯器手動執行一次）
function linkAllMembers() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, COL.userId, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
  let ok = 0, bad = 0;
  ids.forEach(function (id) {
    if (!id) return;
    if (linkMemberMenu_(id)) ok++; else bad++;
    Utilities.sleep(50);
  });
  Logger.log('會員選單：成功 ' + ok + '，失敗 ' + bad);
  return ok;
}

function findMember_(userId) {
  const sheet = getSheet_();
  const row = findRow_(sheet, userId);
  if (!row) return null;
  const v = sheet.getRange(row, 1, 1, HEADERS.length).getDisplayValues()[0];
  return {
    name: v[COL.name - 1],
    phone: v[COL.phone - 1],
    plate: v[COL.plate - 1],
    model: v[COL.model - 1],
    relationship: v[COL.relationship - 1]
  };
}

// ============================================================
// 試算表
// ============================================================
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 手機、車牌設成純文字，避免 09 開頭的 0 被吃掉
    sheet.getRange('A:G').setNumberFormat('@');
    sheet.getRange('H:I').setNumberFormat('yyyy/mm/dd hh:mm');
  }
  // 舊分頁補上新增的欄位（客戶ID、比對結果）
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (head.length < HEADERS.length || head[COL.customerId - 1] !== HEADERS[COL.customerId - 1]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  }
  return sheet;
}

function findRow_(sheet, userId) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const hit = sheet.getRange(2, COL.userId, last - 1, 1)
    .createTextFinder(userId).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : null;
}

// ============================================================
// 工具
// ============================================================
function clean_(v, max) {
  const s = String(v || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, max);
  // 避免被當成試算表公式執行
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function userError_(msg, code) {
  const e = new Error(msg);
  e.userMessage = msg;
  e.code = code || 'INVALID';
  return e;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testMenu() {
  const token = lineToken_();
  Logger.log('token 長度：' + token.length);
  Logger.log('會員選單：' + memberMenuId_(token));
  Logger.log('換選單結果：' + linkMemberMenu_('U9da8ec84b79f48da82d689f8234b3e66'));
}
