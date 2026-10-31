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
  '首次綁定時間', '最後更新時間', '客戶ID', '比對結果', 'Email'
];
const COL = { userId: 1, displayName: 2, name: 3, phone: 4, plate: 5, model: 6, relationship: 7, createdAt: 8, updatedAt: 9, customerId: 10, match: 11, email: 12 };

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
  const t0 = Date.now();
  let action = '', uid = '';
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    action = body.action || '';
    const user = verifyIdToken_(body.idToken);
    uid = user.userId;

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
    // 使用者輸入錯誤（userMessage）不用記；系統錯誤記到「錯誤紀錄」分頁
    if (!err.userMessage) logError_(action, uid, err);
    return json_({ ok: false, error: err.userMessage || '系統忙碌，請稍後再試', code: err.code || 'ERROR' });
  } finally {
    console.log('耗時 ' + action + ' ' + (Date.now() - t0) + ' ms');   // 到「執行項目」點開就看得到每個動作花多久
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

  // 同一個登入憑證 10 分鐘內驗證過就不再問 LINE（不會超過憑證本身的有效期限）
  const cache = CacheService.getScriptCache();
  const key = 'tok_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(idToken)));
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

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
  const user = { userId: data.sub, displayName: data.name || '' };
  const ttl = Math.min(600, Math.floor(Number(data.exp || 0) - Date.now() / 1000) - 30);
  if (ttl > 0) cache.put(key, JSON.stringify(user), ttl);
  return user;
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
    email: String(body.email || '').trim().toLowerCase().slice(0, 100),
    relationship: RELATIONSHIPS.indexOf(body.relationship) >= 0 ? body.relationship : '其他'
  };

  if (!m.name) throw userError_('請填寫姓名');
  if (!/^09\d{8}$/.test(m.phone)) throw userError_('手機號碼格式不正確');
  if (!/^[A-Z0-9]{2,4}-[A-Z0-9]{2,4}$/.test(m.plate)) throw userError_('車牌格式不正確');
  if (m.email && !/^[^\s@=+\-][^\s@]*@[^\s@]+\.[^\s@]+$/.test(m.email)) throw userError_('Email 格式不正確');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw userError_('目前綁定人數較多，請稍候幾秒再按一次', 'BUSY');
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, user.userId);
    const now = new Date();
    const target = row || sheet.getLastRow() + 1;
    const createdAt = row ? sheet.getRange(row, COL.createdAt).getValue() : now;

    // 1) 先把會員本身存進去並立刻存檔：後面任何一步出錯，userId 也一定會留下
    sheet.getRange(target, 1, 1, HEADERS.length).setValues([[
      user.userId, user.displayName, m.name, m.phone, m.plate, m.model, m.relationship,
      createdAt, now, '', '比對中', m.email
    ]]);
    SpreadsheetApp.flush();

    // 2) 換成會員選單（失敗不影響綁定本身）
    let menu = false;
    try { menu = linkMemberMenu_(user.userId); } catch (err) { console.error(err); logError_('bind/換選單', user.userId, err); }

    // 3) 對應旭馳雲端車輛改到背景做（約 1 分鐘內完成），客人不用等
    schedulePending_();

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

// 背景比對：把「比對中」的會員對應到旭馳雲端（綁定後由排程執行；每天凌晨同步時也會補做）
function schedulePending_() {
  try {
    if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processPendingLinks')) return;
    ScriptApp.newTrigger('processPendingLinks').timeBased().after(5 * 1000).create();
  } catch (err) {
    console.error('schedulePending_ ' + err);
    logError_('排程背景比對', '', err);
  }
}

function processPendingLinks() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'processPendingLinks')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  let n = 0;
  rows.forEach(function (v, i) {
    if (v[COL.match - 1] !== '比對中') return;
    const uid = v[COL.userId - 1];
    const m = { phone: v[COL.phone - 1], plate: v[COL.plate - 1], relationship: v[COL.relationship - 1] };
    let link;
    try {
      link = linkToApp_(uid, m);
      SpreadsheetApp.flush();
    } catch (err) {
      console.error(err);
      logError_('背景/比對旭馳雲端', uid, err);
      link = { customerId: '', match: '比對失敗：' + err.message };
    }
    // 寫回前再確認：同一列還是這位會員、資料沒被重新送出或修改過
    const cur = sheet.getRange(i + 2, 1, 1, HEADERS.length).getDisplayValues()[0];
    if (cur[COL.userId - 1] !== uid || cur[COL.match - 1] !== '比對中' || cur[COL.phone - 1] !== m.phone ||
        cur[COL.plate - 1] !== m.plate || cur[COL.relationship - 1] !== m.relationship) return;
    sheet.getRange(i + 2, COL.customerId, 1, 2).setValues([[link.customerId, link.match]]);
    n++;
  });
  SpreadsheetApp.flush();
  Logger.log('背景比對完成 ' + n + ' 筆');
  return n;
}

// ============================================================
// 每天自動補：日後新建檔的車，手機跟已驗證的會員相同 → 自動寫入 LINE
// （先在編輯器執行一次 setupDailySync 建立排程）
// ============================================================
function syncByPhone() {
  requeueNewCars_();       // 綁定時還沒建檔、後來店裡建檔了 → 重新比對
  processPendingLinks();   // 萬一背景排程沒跑到，每天凌晨補做
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

/**
 * 綁定時「找不到車牌（新車待建檔）」、現在旭馳雲端已經有這台車 → 改回「比對中」，
 * 交給 processPendingLinks 用綁定時同一套規則重新比對（車主本人＋手機相同或車子沒填手機才寫入 LINE）
 */
function requeueNewCars_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  const v = SpreadsheetApp.openById(APP_DB_ID).getSheetByName(APP_SHEET).getDataRange().getDisplayValues();
  const cp = v[0].indexOf('車牌');
  if (cp < 0) throw new Error('旭馳雲端「客戶」分頁欄位名稱不符');
  const plates = {};
  for (let i = 1; i < v.length; i++) { const k = normPlate_(v[i][cp]); if (k) plates[k] = true; }
  let n = 0;
  rows.forEach(function (r, i) {
    if (!/^找不到車牌/.test(r[COL.match - 1]) || !plates[normPlate_(r[COL.plate - 1])]) return;
    sheet.getRange(i + 2, COL.match).setValue('比對中');
    n++;
  });
  SpreadsheetApp.flush();
  Logger.log('後來才建檔、重新比對：' + n + ' 筆');
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

// 用別名找出目前的會員選單 ID（選單重建時 menu_live.py alias 會更新別名，這裡不用改）
// fresh = true 時略過快取，直接重查別名
function memberMenuId_(token, fresh) {
  const cache = CacheService.getScriptCache();
  const hit = fresh ? null : cache.get('memberMenuId');
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
  const link = function (menuId) {
    return UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/user/' + encodeURIComponent(userId) + '/richmenu/' + menuId,
      { method: 'post', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
  };
  let res = link(memberMenuId_(token));
  if (res.getResponseCode() !== 200) {
    // 快取的選單 ID 可能已經過時（選單重建過）→ 立刻重查別名再試一次
    res = link(memberMenuId_(token, true));
  }
  if (res.getResponseCode() !== 200) {
    CacheService.getScriptCache().remove('memberMenuId');
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
    email: v[COL.email - 1],
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
  // 舊分頁補上新增的欄位（客戶ID、比對結果、Email）；檢查過一次就記起來，不用每次都讀標題
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('HEADERS_OK') !== String(HEADERS.length)) {
    const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (head.length < HEADERS.length || head[HEADERS.length - 1] !== HEADERS[HEADERS.length - 1]) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    }
    props.setProperty('HEADERS_OK', String(HEADERS.length));
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

// 系統錯誤寫進「錯誤紀錄」分頁（匿名呼叫時 Google 不保留錯誤內容，只能自己記）
function logError_(where, userId, err) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('錯誤紀錄');
    if (!sh) {
      sh = ss.insertSheet('錯誤紀錄');
      sh.appendRow(['時間', '位置', 'LINE UserID', '錯誤訊息', '詳細']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), where, userId || '', String((err && err.message) || err), String((err && err.stack) || '').slice(0, 1000)]);
    SpreadsheetApp.flush();
  } catch (e) {
    console.error('logError_ 失敗', e);
  }
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
