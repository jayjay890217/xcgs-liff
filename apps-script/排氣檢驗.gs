/**
 * 旭馳車業｜檢驗排氣通知（和 Code.gs 放在同一個 Apps Script 專案「liff」）
 *
 * 每月 15 號早上 9 點：
 *   1) 把有綁 LINE 的車（旭馳APP資料庫「客戶」分頁 LINE_User_ID 有值）拿去環境部查檢驗紀錄，結果寫進「排氣檢驗」分頁
 *      ‧ 已經知道發照月、跟這個月無關的車不重查；「這個月發照」「已逾期」「還沒查過／查不到」的才查
 *   2) 「這個月是發照月、還沒驗」和「已逾期、還沒驗」的車 → 用客人的官方帳號 LINE 推播「檢驗排氣通知」
 *      逾期的每個月 15 號都會再提醒，直到查到合格紀錄
 *   3) 推播完把清單傳到員工 LINE（員工名單「備註」有寫「收通知」的人）
 * 一次最多跑 5 分鐘，車多查不完會自動接著跑，全部查完才推播。
 *
 * 第一次使用：在編輯器執行 setupEpaMonthly（授權＋建立每月排程＋先把全部車查一遍，這次不推播）
 * 員工 LINE 清單：指令碼屬性 STAFF_NOTIFY_KEY 要填員工綁定程式的 WEBHOOK_KEY（老闆自己貼，沒填就只寫試算表）
 * 其他手動工具：epaScanAll（全部重查、不推播）、epaPreview（看「現在跑的話會通知誰」，不推播）、epaRunNow（立刻照 15 號流程跑一次）
 */
const EPA_URL = 'https://mobile.moenv.gov.tw/Motor/query/Query_Check.aspx';
const EPA_TZ = 'Asia/Taipei';
const EPA_SHEET = '排氣檢驗';
const EPA_HEADERS = ['車牌', '客戶ID', 'LINE UserID', '出廠年月', '發照月', '驗車期間', '狀態', '說明', '最後檢驗', '查詢時間', '上次通知', '通知次數'];
const EC = { plate: 1, customerId: 2, userId: 3, made: 4, licMonth: 5, window: 6, status: 7, msg: 8, last: 9, checkedAt: 10, notifiedAt: 11, count: 12 };
const EPA_STATUS = { done: '本期已驗', notyet: '還沒到期間', young: '未滿5年免驗', due: '驗車期間未驗', overdue: '已逾期', unknown: '查無車籍', error: '查詢失敗' };
const EPA_BUDGET_MS = 5 * 60 * 1000;    // Apps Script 一次最多 6 分鐘，留 1 分鐘寫結果
const EPA_DAY = 15, EPA_HOUR = 9;

// 員工 LINE（旭馳員工綁定）網址
const STAFF_NOTIFY_URL = 'https://script.google.com/macros/s/AKfycbyB4WHZhAAKkXlpJADqooRXtjiOtG92FCNq-a7lMjb5TJHR7wiJJfMK_k1qX7QO6mY/exec';

// 客人收到的卡片：標題、每台車一行、說明；每台車一個「此台車不是我的」按鈕 → 會員頁解除這台車並重新綁定
const MEMBER_LIFF_URL = 'https://liff.line.me/2011218129-CeYpPJHj';
const EPA_NOTE = '機車出廠滿 5 年每年都要做排氣定期檢驗，期限是發照月份的前後一個月，逾期未驗可能會被開罰喔！已經驗過的話請忽略這則訊息 🙏';

// ============================================================
// 排程入口
// ============================================================
function setupEpaMonthly() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'epaMonthly') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('epaMonthly').timeBased().onMonthDay(EPA_DAY).atHour(EPA_HOUR).inTimezone(EPA_TZ).create();
  Logger.log('已建立每月 ' + EPA_DAY + ' 號 ' + EPA_HOUR + ' 點自動檢查排氣檢驗');
  epaScanAll();
}

/** 每月 15 號排程 */
function epaMonthly() { epaStart_('month'); }
/** 手動：立刻照 15 號流程跑一次（會推播！） */
function epaRunNow() { epaStart_('month'); }
/** 手動：全部車重查一遍，只寫試算表、不推播 */
function epaScanAll() { epaStart_('all'); }

function epaStart_(mode) {
  const job = { mode: mode, start: Date.now(), month: Utilities.formatDate(new Date(), EPA_TZ, 'yyyy-MM') };
  PropertiesService.getScriptProperties().setProperty('EPA_JOB', JSON.stringify(job));
  epaContinue();
}

/** 查一批；沒查完就排 1 分鐘後接著跑，查完了（月排程）才推播 */
function epaContinue() {
  const t0 = Date.now();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'epaContinue').forEach(t => ScriptApp.deleteTrigger(t));
  const props = PropertiesService.getScriptProperties();
  const job = JSON.parse(props.getProperty('EPA_JOB') || 'null');
  if (!job) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { epaLater_(); return; }
  try {
    const sh = epaSheet_();
    const rows = epaSyncCars_(sh);
    const now = new Date(), tm = Number(Utilities.formatDate(now, EPA_TZ, 'M'));
    const since = Math.floor(job.start / 60000) * 60000;    // 查詢時間只記到分鐘
    const todo = rows.filter(r => epaCheckedMs_(r) < since && (job.mode === 'all' || epaNeedQuery_(r, tm)));
    let n = 0;
    for (const r of todo) {
      if (Date.now() - t0 > EPA_BUDGET_MS) break;
      epaWriteResult_(sh, r, epaQuery_(r.plate, now));
      n++;
      Utilities.sleep(300);   // 別對環境部網站連續狂查
    }
    SpreadsheetApp.flush();
    Logger.log('排氣檢驗：這次查了 ' + n + ' 台，還剩 ' + (todo.length - n) + ' 台');
    props.deleteProperty('EPA_FAILS');
    if (n < todo.length) { epaLater_(); return; }

    props.deleteProperty('EPA_JOB');
    if (job.mode === 'month') epaSend_(sh, job.month);
  } catch (err) {
    console.error(err);
    logError_('排氣檢驗', '', err);
    const fails = Number(props.getProperty('EPA_FAILS') || 0);
    if (fails < 3) {      // 暫時出錯 → 1 分鐘後再試，最多 3 次
      props.setProperty('EPA_FAILS', String(fails + 1));
      epaLater_();
    } else {
      props.deleteProperty('EPA_FAILS');
      props.deleteProperty('EPA_JOB');
      epaStaffNotify_('⚠️ 檢驗排氣通知：自動查詢失敗（' + err.message + '），請找 Claude 看會員綁定試算表「錯誤紀錄」分頁。');
    }
  } finally {
    lock.releaseLock();
  }
}

function epaLater_() {
  ScriptApp.newTrigger('epaContinue').timeBased().after(60 * 1000).create();
}

/** 哪些車在月排程要重查：沒查過／查不到／這個月發照／已經在期間或逾期（看驗了沒） */
function epaNeedQuery_(r, tm) {
  if (!r.licMonth) return true;
  if (['due', 'overdue'].indexOf(r.statusKey) >= 0) return true;
  return Number(r.licMonth) === tm;
}

// ============================================================
// 推播
// ============================================================
/** 這次要通知哪些車：發照月（含下個月、期間最後一個月）還沒驗 ＋ 已逾期還沒驗（逾期的每個月都提醒）；發照月前一個月先不通知 */
function epaTargets_(rows, tm) {
  return rows.filter(r => r.userId && (r.statusKey === 'overdue' ||
    (r.statusKey === 'due' && (tm - Number(r.licMonth) + 12) % 12 <= 1)));
}

function epaSend_(sh, month) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('EPA_SENT') === month) { Logger.log(month + ' 已經推播過，不重複發'); return; }
  props.setProperty('EPA_SENT', month);

  const rows = epaRows_(sh).filter(r => r.userId);
  const tm = Number(month.slice(5));
  const targets = epaTargets_(rows, tm);
  const names = epaMemberNames_();
  const byUser = {};
  targets.forEach(r => (byUser[r.userId] = byUser[r.userId] || []).push(r));

  const today = Utilities.formatDate(new Date(), EPA_TZ, 'yyyy/MM/dd');
  const report = [];
  let ok = 0, bad = 0;
  Object.keys(byUser).forEach(uid => {
    const cars = byUser[uid];
    const sent = epaPush_(uid, epaCard_(names[uid], cars));
    cars.forEach(r => {
      const nth = Number(r.count || 0) + 1;
      if (sent) sh.getRange(r.row, EC.notifiedAt, 1, 2).setValues([[today, String(nth)]]);
      report.push((sent ? '' : '❌ ') + r.plate + ' ' + (names[uid] || '') + ' ' +
        (r.statusKey === 'overdue' ? '逾期（第 ' + nth + ' 次提醒）' : '驗車期間未驗（' + r.window + '）'));
    });
    sent ? ok++ : bad++;
    Utilities.sleep(100);
  });
  SpreadsheetApp.flush();

  const fail = rows.filter(r => r.statusKey === 'error').length;
  epaStaffNotify_('【檢驗排氣通知】' + today + ' 已發送\n通知 ' + ok + ' 位客人、' + targets.length + ' 台車'
    + (bad ? '（' + bad + ' 位發送失敗，可能已封鎖官方帳號）' : '') + '\n\n'
    + (report.length ? report.join('\n') : '這個月沒有需要通知的車')
    + (fail ? '\n\n查詢失敗 ' + fail + ' 台，看會員綁定試算表「排氣檢驗」分頁' : ''));
  Logger.log('推播完成：' + ok + ' 位成功、' + bad + ' 位失敗');
}

function epaCarLine_(r) {
  return '🛵 ' + r.plate + '：' + (r.statusKey === 'overdue'
    ? '已經逾期（應在 ' + r.window + ' 驗車），目前查不到合格紀錄，請盡快檢驗'
    : '現在是驗車期間（' + r.window + '），目前還沒有檢驗紀錄，請在期限內完成檢驗');
}

function epaCard_(name, cars) {
  const t = (text, o) => Object.assign({ type: 'text', text: text, wrap: true, size: 'sm', color: '#333333' }, o || {});
  const body = [t((name ? name + ' ' : '') + '您好，旭馳車業提醒您：')];
  cars.forEach(r => body.push(t(epaCarLine_(r), { margin: 'md', weight: 'bold' })));
  body.push(t(EPA_NOTE, { margin: 'lg', size: 'xs', color: '#888888' }));
  const foot = [t('車子已經賣掉或換車了？按下面的按鈕，再輸入您現在的車牌', { size: 'xs', color: '#888888' })];
  cars.forEach(r => foot.push({ type: 'button', style: 'secondary', height: 'sm', action: { type: 'uri',
    label: cars.length > 1 ? r.plate + ' 不是我的車' : '此台車不是我的', uri: MEMBER_LIFF_URL + '?notmine=' + encodeURIComponent(r.plate) } }));
  return {
    type: 'flex', altText: '【檢驗排氣通知】' + cars.map(r => r.plate).join('、') + ' 排氣檢驗提醒',
    contents: {
      type: 'bubble',
      header: { type: 'box', layout: 'vertical', backgroundColor: '#E8590C',
        contents: [t('檢驗排氣通知', { weight: 'bold', size: 'lg', color: '#FFFFFF' })] },
      body: { type: 'box', layout: 'vertical', contents: body },
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: foot }
    }
  };
}

function epaPush_(userId, message) {
  const token = lineToken_();
  if (!token) throw new Error('沒有設定 LINE_CHANNEL_ACCESS_TOKEN');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ to: userId, messages: [message] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) return true;
  logError_('排氣通知推播', userId, new Error('HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200)));
  return false;
}

/** 傳到員工 LINE（員工綁定程式 notifyBoss → 備註有「收通知」的人） */
function epaStaffNotify_(text) {
  const key = PropertiesService.getScriptProperties().getProperty('STAFF_NOTIFY_KEY');
  if (!key) { Logger.log('沒設定 STAFF_NOTIFY_KEY，略過員工 LINE：\n' + text); return; }
  try {
    const res = UrlFetchApp.fetch(STAFF_NOTIFY_URL, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ action: 'notifyBoss', key: key, text: text })
    });
    Logger.log('員工 LINE：' + res.getContentText().slice(0, 200));
  } catch (err) {
    logError_('排氣通知/員工LINE', '', err);
  }
}

/** 手動：看「現在跑的話會通知誰」（只看試算表現有資料，不查、不推播） */
function epaPreview() {
  const tm = Number(Utilities.formatDate(new Date(), EPA_TZ, 'M'));
  const names = epaMemberNames_();
  const t = epaTargets_(epaRows_(epaSheet_()), tm);
  Logger.log('會通知 ' + t.length + ' 台：\n' + t.map(r => (names[r.userId] || '') + '｜' + epaCarLine_(r)).join('\n'));
  if (t.length) Logger.log('卡片範例：\n' + JSON.stringify(epaCard_(names[t[0].userId], [t[0]])));
}

// ============================================================
// 試算表「排氣檢驗」
// ============================================================
function epaSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(EPA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(EPA_SHEET);
    sh.getRange(1, 1, 1, EPA_HEADERS.length).setValues([EPA_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:L').setNumberFormat('@');
  }
  return sh;
}

function epaRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const statusKey = {};
  Object.keys(EPA_STATUS).forEach(k => statusKey[EPA_STATUS[k]] = k);
  return sh.getRange(2, 1, last - 1, EPA_HEADERS.length).getDisplayValues().map((v, i) => ({
    row: i + 2, plate: v[EC.plate - 1], customerId: v[EC.customerId - 1], userId: v[EC.userId - 1],
    licMonth: v[EC.licMonth - 1], window: v[EC.window - 1], statusKey: statusKey[v[EC.status - 1]] || '',
    checkedAt: v[EC.checkedAt - 1], count: v[EC.count - 1]
  }));
}

function epaCheckedMs_(r) {
  const m = String(r.checkedAt || '').match(/^(\d{4})-(\d\d)-(\d\d) (\d\d):(\d\d)$/);
  return m ? new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':00+08:00').getTime() : 0;
}

/** 從旭馳APP資料庫抓有綁 LINE 的車，新車加進分頁、LINE 有變就更新；解除綁定的車清掉 LINE（不再通知） */
function epaSyncCars_(sh) {
  const v = SpreadsheetApp.openById(APP_DB_ID).getSheetByName(APP_SHEET).getDataRange().getDisplayValues();
  const h = v[0], c = { id: h.indexOf('客戶ID'), plate: h.indexOf('車牌'), line: h.indexOf('LINE_User_ID') };
  if (c.plate < 0 || c.line < 0) throw new Error('旭馳雲端「客戶」分頁欄位名稱不符');
  const cars = {};
  for (let i = 1; i < v.length; i++) {
    const plate = epaPlate_(v[i][c.plate]), uid = v[i][c.line];
    if (plate && uid && !cars[normPlate_(plate)]) cars[normPlate_(plate)] = { plate: plate, id: c.id >= 0 ? v[i][c.id] : '', uid: uid };
  }
  const have = {};
  epaRows_(sh).forEach(r => {
    const k = normPlate_(r.plate), car = cars[k];
    have[k] = true;
    const uid = car ? car.uid : '';
    if (uid !== r.userId) sh.getRange(r.row, EC.userId).setValue(uid);
  });
  const add = Object.keys(cars).filter(k => !have[k]).map(k => {
    const row = EPA_HEADERS.map(() => '');
    row[EC.plate - 1] = cars[k].plate; row[EC.customerId - 1] = cars[k].id; row[EC.userId - 1] = cars[k].uid;
    return row;
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, EPA_HEADERS.length).setValues(add);
  SpreadsheetApp.flush();
  return epaRows_(sh).filter(r => r.userId);
}

function epaWriteResult_(sh, r, x) {
  const at = Utilities.formatDate(new Date(), EPA_TZ, 'yyyy-MM-dd HH:mm');
  sh.getRange(r.row, EC.made, 1, 7).setValues([[
    x.made || '', x.licMonth ? String(x.licMonth) : (r.licMonth || ''), x.window || '',
    EPA_STATUS[x.status] || x.status, x.msg || '', x.last || '', at
  ]]);
  r.statusKey = x.status; r.licMonth = x.licMonth || r.licMonth; r.window = x.window || r.window;
}

/** 查一台；網站錯誤不中斷整批，記成「查詢失敗」下次再查 */
function epaQuery_(plate, now) {
  try {
    return epaJudge_(plate, epaFetch_(plate), now);
  } catch (err) {
    return { status: 'error', msg: String(err.message || err).slice(0, 200) };
  }
}

/** 會員綁定分頁：LINE UserID → 姓名 */
function epaMemberNames_() {
  const sh = getSheet_(), last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, COL.name).getDisplayValues().forEach(r => { if (r[0]) out[r[0]] = r[COL.name - 1]; });
  return out;
}

/** 車牌整理成環境部要的格式（ABC-1234） */
function epaPlate_(s) {
  const p = String(s || '').toUpperCase().replace(/[^A-Z0-9-]/g, '');
  if (!p || p.indexOf('-') >= 0) return p;
  const m = p.match(/^([A-Z]+)(\d+)$/) || p.match(/^(\d+)([A-Z]+)$/);
  return m ? m[1] + '-' + m[2] : p;
}

// ============================================================
// 環境部查詢（和工作台 KEY單 的排氣檢驗同一套判斷）
// 出廠滿 5 年每年要驗，期限＝發照月份的前一個月～後一個月
// ============================================================
function epaFetch_(plate) {
  let cookie = '';
  const req = (payload) => {
    const opt = { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } };
    if (cookie) opt.headers.Cookie = cookie;
    if (payload) { opt.method = 'post'; opt.payload = payload; }
    const res = UrlFetchApp.fetch(EPA_URL, opt);
    const sc = res.getAllHeaders()['Set-Cookie'];
    if (sc) cookie = [].concat(sc).map(c => String(c).split(';')[0]).join('; ');
    if (res.getResponseCode() !== 200) throw new Error('環境部網站沒有回應（' + res.getResponseCode() + '）');
    return res.getContentText('UTF-8');
  };
  const hidden = (h) => {
    const o = {};
    ['__VIEWSTATE', '__VIEWSTATEGENERATOR', '__EVENTVALIDATION'].forEach(n => {
      const m = h.match(new RegExp('id="' + n + '" value="([^"]*)"'));
      o[n] = m ? m[1] : '';
    });
    return o;
  };
  const p1 = req();
  const p2 = req(Object.assign(hidden(p1), { 'ctl00$MainContent$txtCarNo': plate, 'ctl00$MainContent$btnQuery': '查詢' }));
  let html = p2;
  if (p2.indexOf('lbuttonMoreData') >= 0) {
    try {
      html = req(Object.assign(hidden(p2), { '__EVENTTARGET': 'ctl00$MainContent$lbuttonMoreData', '__EVENTARGUMENT': '',
        'ctl00$MainContent$txtCarNo': plate }));
    } catch (e) { html = p2; }
  }
  return html;
}

function epaTables_(html) {
  const cell = s => s.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
  return (html.match(/<table[\s\S]*?<\/table>/g) || []).map(t =>
    (t.match(/<tr[\s\S]*?<\/tr>/g) || []).map(r => (r.match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/g) || []).map(cell)));
}

/** status：done 今年已驗／due 現在該驗／overdue 已逾期／notyet 還沒到驗車期間／young 未滿 5 年免驗／unknown 查無車籍 */
function epaJudge_(plate, html, now) {
  const tabs = epaTables_(html);
  const find = w => tabs.filter(t => t.length && t[0].indexOf(w) >= 0)[0] || [];
  const tt = find('檢驗別'), ct = find('發照日期');
  const th = tt[0] || [], ch = ct[0] || [];
  const tests = tt.slice(1).filter(r => r.length >= th.length && r[th.indexOf('檢測日期')]).map(r => ({
    kind: r[th.indexOf('檢驗別')], result: r[th.indexOf('檢測結果')], date: r[th.indexOf('檢測日期')],
    year: r[th.indexOf('檢測標籤/年度')] || ''
  }));
  const car = (ct.slice(1).filter(r => r.length >= ch.length && /\d{6}/.test(r[ch.indexOf('出廠日')] || ''))[0]) || null;
  const out = { plate, status: 'unknown', need: false };
  if (!car) { out.msg = '環境部查無車籍資料（新領牌、過戶或報廢的車可能查不到）'; return out; }

  const made = car[ch.indexOf('出廠日')], lic = car[ch.indexOf('發照日期')];
  const fy = Number(made.slice(0, 4)), fm = Number(made.slice(4, 6)), lm = Number(lic.slice(4, 6));
  Object.assign(out, { made: fy + '/' + ('0' + fm).slice(-2), licMonth: lm });
  const mon = (y, m) => { const d = new Date(Date.UTC(y, m - 1, 1)); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1 }; };
  const key = o => o.y * 100 + o.m;
  const win = y => ({ from: mon(y, lm - 1), to: mon(y, lm + 1) });
  const fmt = w => w.from.y + '/' + w.from.m + '～' + w.to.y + '/' + w.to.m;
  const must = y => (win(y).to.y * 12 + win(y).to.m) - (fy * 12 + fm) >= 60;   // 驗車期間結束前滿 5 年
  const done = y => tests.some(t => t.kind.indexOf('定期') >= 0 && /合格/.test(t.result) && t.result.indexOf('不合格') < 0 &&
    Number(t.year) === y - 1911);
  now = now || new Date();
  const ty = Number(Utilities.formatDate(now, EPA_TZ, 'yyyy')), tm = Number(Utilities.formatDate(now, EPA_TZ, 'M'));
  const today = ty * 100 + tm;

  let y = ty;
  if (today >= key(win(ty + 1).from)) y = ty + 1;       // 1 月發照的車，12 月就進入明年那一期
  else if (today < key(win(ty).from)) y = ty - 1;      // 今年的驗車期間還沒到 → 先看去年那一期
  const w = win(y);
  out.window = fmt(w);
  const last = tests.filter(t => t.kind.indexOf('定期') >= 0)[0];
  if (last) out.last = last.date.slice(0, 4) + '/' + last.date.slice(4, 6) + '/' + last.date.slice(6, 8) + ' ' + last.result;

  if (!must(y)) {
    let ny = y + 1; while (!must(ny) && ny < ty + 6) ny++;
    out.status = 'young'; out.window = fmt(win(ny));
    out.msg = '出廠未滿 5 年，免驗；第一次要驗：' + out.window;
  } else if (done(y)) {
    out.status = y >= ty ? 'done' : 'notyet';
    out.window = fmt(win(y >= ty ? y + 1 : ty));
    out.msg = (y >= ty ? '這一期已驗過' : '去年已驗，今年還沒到期間') + '；下次驗車：' + out.window;
  } else if (today <= key(w.to)) {
    out.status = 'due'; out.need = true;
    out.msg = '現在就是驗車期間（' + out.window + '），還沒驗';
  } else {
    out.status = 'overdue'; out.need = true;
    out.msg = '已經逾期！應在 ' + out.window + ' 驗車，目前查不到合格紀錄';
  }
  return out;
}

// ============================================================
// 「此台車不是我的」：客人從卡片按鈕打開會員頁 → 會員頁呼叫 doPost action notMine
// 解除旭馳雲端這台車的 LINE、記到「不是我的車」分頁（每日同手機自動補 LINE 會跳過）、通知員工 LINE
// ============================================================
const NOT_MINE_SHEET = '不是我的車';

function notMine_(user, body) {
  const plate = epaPlate_(body.plate), key = normPlate_(plate);
  if (!key) throw userError_('缺少車牌');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw userError_('目前人數較多，請稍候幾秒再試', 'BUSY');
  try {
    let n = 0;
    const sh = SpreadsheetApp.openById(APP_DB_ID).getSheetByName(APP_SHEET);
    const v = sh.getDataRange().getDisplayValues();
    const cp = v[0].indexOf('車牌'), cl = v[0].indexOf('LINE_User_ID');
    for (let i = 1; i < v.length; i++) {
      if (normPlate_(v[i][cp]) === key && v[i][cl] === user.userId) { sh.getRange(i + 1, cl + 1).setValue(''); n++; }
    }
    const es = epaSheet_();
    epaRows_(es).forEach(r => { if (normPlate_(r.plate) === key && r.userId === user.userId) es.getRange(r.row, EC.userId).setValue(''); });

    const ms = getSheet_(), row = findRow_(ms, user.userId);
    const name = row ? ms.getRange(row, COL.name).getDisplayValue() : '';
    if (row && normPlate_(ms.getRange(row, COL.plate).getDisplayValue()) === key) ms.getRange(row, COL.match).setValue('客人表示不是他的車，待重新綁定');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = ss.getSheetByName(NOT_MINE_SHEET);
    if (!log) { log = ss.insertSheet(NOT_MINE_SHEET); log.appendRow(['時間', 'LINE UserID', '姓名', '車牌', '解除台數']); log.setFrozenRows(1); }
    if (!notMineSet_()[user.userId + '|' + key]) log.appendRow([new Date(), user.userId, name, plate, n]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('notMine');

    epaStaffNotify_('🚫 客人 ' + (name || user.displayName || '') + ' 表示 ' + plate + ' 不是他的車，已解除這台車的 LINE（旭馳雲端這台車的車主電話可能要更新）');
    return { ok: true, plate: plate, removed: n, member: findMember_(user.userId) };
  } finally {
    lock.releaseLock();
  }
}

/** 客人說過「不是我的」的車：{ 'LINE UserID|車牌': true }（同手機自動補 LINE 時跳過） */
function notMineSet_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('notMine');
  if (hit) return JSON.parse(hit);
  const out = {}, sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOT_MINE_SHEET);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 2, sh.getLastRow() - 1, 3).getDisplayValues().forEach(r => { out[r[0] + '|' + normPlate_(r[2])] = true; });
  try { cache.put('notMine', JSON.stringify(out), 600); } catch (e) {}
  return out;
}

/** 在編輯器執行：測試查詢（不寫試算表、不推播） */
function testEpaNotify() {
  ['325-KGH', 'MPA-3988'].forEach(p => Logger.log(p + ' → ' + JSON.stringify(epaQuery_(p))));
}
